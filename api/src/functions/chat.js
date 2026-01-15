import { app } from "@azure/functions";

// Placeholder chat route.
// Replace this with Azure OpenAI or OpenAI API call once you configure secrets.
app.http("chat", {
  route: "chat",
  methods: ["POST"],
  authLevel: "anonymous",
  handler: async (request) => {
    const body = await request.json().catch(() => ({}));
    const last = Array.isArray(body?.messages) ? body.messages[body.messages.length - 1] : null;
    const userText = last?.content ? String(last.content) : "";

    return {
      status: 200,
      jsonBody: {
        answer:
          "Chat backend placeholder is running. Configure Azure OpenAI/OpenAI and replace api/src/functions/chat.js.\n\nYour last message was:\n" +
          userText,
      },
    };
  },
});
