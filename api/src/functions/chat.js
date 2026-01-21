import { app } from "@azure/functions";

app.http("chat", {
  methods: ["POST"],
  authLevel: "anonymous",
  handler: async (req) => {
    try {
      const { messages, context: appContext } = await req.json();

      const apiKey = process.env.OPENAI_API_KEY;
      const model = process.env.OPENAI_MODEL || "gpt-4o-mini";

      if (!apiKey) {
        return {
          status: 500,
          jsonBody: {
            error:
              "OPENAI_API_KEY is not set. Add it in Azure Static Web Apps > Configuration > Application settings.",
          },
        };
      }

      const systemText =
        "You are a helpful assistant for a design programme tracker web app. " +
        "Use the provided APP_CONTEXT_JSON to answer questions accurately. " +
        "If something is missing, say what you’d need.";

      const input = [
        { role: "system", content: systemText },
        { role: "system", content: `APP_CONTEXT_JSON:\n${JSON.stringify(appContext ?? {}, null, 2)}` },
        ...(Array.isArray(messages) ? messages : []),
      ];

      const r = await fetch("https://api.openai.com/v1/responses", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          Authorization: `Bearer ${apiKey}`,
        },
        body: JSON.stringify({
          model,
          input,
          temperature: 0.2,
          max_output_tokens: 800,
        }),
      });

      const data = await r.json().catch(() => ({}));

      if (!r.ok) {
        return {
          status: 500,
          jsonBody: {
            error: "OpenAI request failed",
            status: r.status,
            details: data,
          },
        };
      }

      const answer =
        data?.output_text ||
        data?.output?.[0]?.content?.[0]?.text ||
        "No response text returned.";

      return {
        status: 200,
        jsonBody: { answer },
      };
    } catch (e) {
      return {
        status: 500,
        jsonBody: { error: "chat handler error", details: String(e) },
      };
    }
  },
});
