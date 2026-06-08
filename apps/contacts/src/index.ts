import { run } from "@microsoft-mcp/core";
import { tools } from "./tools.js";

void run({ name: "microsoft-contacts", version: "1.0.0", title: "Microsoft Contacts", scopes: ["User.Read", "Contacts.ReadWrite"] }, tools);
