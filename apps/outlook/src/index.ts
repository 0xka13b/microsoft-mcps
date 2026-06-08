import { run } from "@microsoft-mcp/core";
import { tools } from "./tools.js";

void run({ name: "microsoft-outlook", version: "1.0.0", title: "Microsoft Outlook", scopes: ["User.Read", "Mail.ReadWrite", "Mail.Send"] }, tools);
