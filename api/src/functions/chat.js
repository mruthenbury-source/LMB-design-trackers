export async function handler(req) {
  try {
    const { messages, context: appContext } = await req.json();

    const apiKey = process.env.OPENAI_API_KEY;
    const model = process.env.OPENAI_MODEL || "gpt-4o-mini";

    if (!apiKey) {
      return new Response(
        JSON.stringify({
          error:
            "OPENAI_API_KEY is not set. Add it in Azure Static Web Apps > Configuration > Application settings.",
        }),
        { status: 500, headers: { "Content-Type": "application/json" } }
      );
    }

    // Build a compact, safe system instruction + embed your app context.
    const systemText =
      "You are a helpful assistant for a design programme tracker web app. " +
      "Use the provided APP_CONTEXT_JSON to answer questions accurately. " +
      "If the user asks for something not in the context, say what you’d need.";

    const input = [
      { role: "system", content: systemText },
      {
        role: "system",
        content: `APP_CONTEXT_JSON:\n${JSON.stringify(appContext ?? {}, null, 2)}`,
      },
      ...(Array.isArray(messages) ? messages : []),
    ];

    const r = await fetch("https://api.openai.com/v1/responses", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${apiKey}`, // Bearer auth per OpenAI docs :contentReference[oaicite:3]{index=3}
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
      return new Response(
        JSON.stringify({
          error: "OpenAI request failed",
          status: r.status,
          details: data,
        }),
        { status: 500, headers: { "Content-Type": "application/json" } }
      );
    }

    // Responses API returns text in output_text
    const answer =
      data?.output_text ||
      data?.output?.[0]?.content?.[0]?.text ||
      "No response text returned.";

    return new Response(JSON.stringify({ answer }), {
      headers: { "Content-Type": "application/json" },
    });
  } catch (e) {
    return new Response(
      JSON.stringify({ error: "chat handler error", details: String(e) }),
      { status: 500, headers: { "Content-Type": "application/json" } }
    );
  }
}

