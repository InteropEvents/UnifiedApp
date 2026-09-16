import axios from 'axios';

type ChatCompletionResponse = {
  choices?: Array<{
    message?: {
      content?: string;
    };
  }>;
};

export function getAzureOpenAIChatCompletionsUrl(): string {
  const resourceName = process.env.REACT_APP_OPENAI_RES_NAME?.trim();
  const deploymentId = process.env.REACT_APP_OPENAI_DEPLOY_ID?.trim();
  const apiVersion = process.env.REACT_APP_OPENAI_API_VERSION?.trim() || '2024-10-21';

  if (!resourceName || !deploymentId) {
    throw new Error('Azure OpenAI resource name or deployment ID is not configured.');
  }

  return `https://${resourceName}.openai.azure.com/openai/deployments/${encodeURIComponent(deploymentId)}/chat/completions?api-version=${encodeURIComponent(apiVersion)}`;
}

export async function summarizeMeetingTranscript(
  meetingSubject: string,
  transcript: string
): Promise<string> {
  const apiKey = process.env.REACT_APP_OPENAI_API_KEY?.trim();

  if (!apiKey) {
    throw new Error('Azure OpenAI API key is not configured.');
  }

  const response = await axios.post<ChatCompletionResponse>(
    getAzureOpenAIChatCompletionsUrl(),
    {
      messages: [
        {
          role: 'system',
          content: 'You create concise meeting summaries. Return sections named Summary, Decisions, Action items, and Open questions. Use bullets, preserve stated owners and dates, and do not invent missing details.',
        },
        {
          role: 'user',
          content: `Meeting: ${meetingSubject}\n\nSummarize this synthetic demo transcript:\n\n${transcript}`,
        },
      ],
      temperature: 0.2,
      max_tokens: 800,
    },
    {
      headers: {
        'Content-Type': 'application/json',
        'api-key': apiKey,
      },
    }
  );

  const summary = response.data.choices?.[0]?.message?.content?.trim();
  if (!summary) {
    throw new Error('Azure OpenAI returned an empty meeting summary.');
  }

  return summary;
}