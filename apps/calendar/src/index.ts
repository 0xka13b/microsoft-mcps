import { run } from "@microsoft-mcp/core";
import { tools } from "./tools.js";

void run({ name: "microsoft-calendar", version: "1.0.0", title: "Microsoft Calendar", scopes: ["User.Read", "Calendars.ReadWrite"] }, tools);
