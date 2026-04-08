import { config } from "dotenv";

config();

const required = ["MicrosoftAppId", "MicrosoftAppPassword"];

for (const key of required) {
  if (!process.env[key]) {
    process.env[key] = "";
  }
}

export const env = {
  port: Number(process.env.PORT ?? 3978),
  microsoftAppId: process.env.MicrosoftAppId ?? "",
  microsoftAppPassword: process.env.MicrosoftAppPassword ?? "",
  nodeEnv: process.env.NODE_ENV ?? "development"
};
