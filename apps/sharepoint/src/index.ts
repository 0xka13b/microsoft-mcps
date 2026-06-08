import { run } from "@microsoft-mcp/core";
import { tools } from "./tools.js";

void run({ name: "microsoft-sharepoint", version: "1.0.0", title: "Microsoft SharePoint", scopes: ["User.Read", "Sites.ReadWrite.All"] }, tools);
