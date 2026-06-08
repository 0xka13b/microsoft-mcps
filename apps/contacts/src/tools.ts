import { z } from "zod";
import { defineTool } from "@microsoft-mcp/core";
import { escapeKql, validateId } from "@microsoft-mcp/validation";

export const tools = [
  defineTool({
    name: "me",
    description: "Get the signed-in user's profile (id, displayName, mail).",
    inputSchema: {},
    confirmationPolicy: "never",
    handler: ({ graph }) =>
      graph.request("GET", "/me", undefined, {
        $select: "id,displayName,mail",
      }),
  }),

  defineTool({
    name: "list_contacts",
    description: "List contacts from the user's contact folder.",
    inputSchema: {
      limit: z
        .number()
        .int()
        .min(1)
        .max(100)
        .optional()
        .describe("Max contacts to return. Defaults to 50."),
    },
    confirmationPolicy: "never",
    handler: async ({ graph }, params) => {
      const data = await graph.request("GET", "/me/contacts", undefined, {
        $top: String(params.limit ?? 50),
        $select:
          "id,displayName,givenName,surname,emailAddresses,businessPhones,mobilePhone,jobTitle,companyName",
      });
      return data.value ?? [];
    },
  }),

  defineTool({
    name: "get_contact",
    description: "Get a single contact by ID.",
    inputSchema: {
      contact_id: z.string().describe("Contact ID"),
    },
    confirmationPolicy: "never",
    handler: ({ graph }, params) => {
      validateId(params.contact_id, "contact_id");
      return graph.request("GET", `/me/contacts/${params.contact_id}`);
    },
  }),

  defineTool({
    name: "create_contact",
    description: "Create a new contact.",
    inputSchema: {
      given_name: z.string().describe("First name"),
      surname: z.string().optional().describe("Last name"),
      email_addresses: z
        .array(
          z.object({
            address: z.string().describe("Email address"),
            name: z.string().optional().describe("Display name for this email"),
          }),
        )
        .optional()
        .describe("Email addresses"),
      phone_numbers: z
        .array(
          z.object({
            type: z
              .enum(["home", "business", "mobile", "other"])
              .describe("Phone type"),
            number: z.string().describe("Phone number"),
          }),
        )
        .optional()
        .describe("Phone numbers"),
      job_title: z.string().optional().describe("Job title"),
      company_name: z.string().optional().describe("Company name"),
    },
    confirmationPolicy: "always",
    handler: ({ graph }, params) => {
      const contact: any = { givenName: params.given_name };

      if (params.surname) contact.surname = params.surname;
      if (params.job_title) contact.jobTitle = params.job_title;
      if (params.company_name) contact.companyName = params.company_name;

      if (params.email_addresses?.length) {
        contact.emailAddresses = params.email_addresses.map((e) => ({
          address: e.address,
          name: e.name ?? e.address,
        }));
      }

      if (params.phone_numbers?.length) {
        for (const phone of params.phone_numbers) {
          if (phone.type === "mobile") contact.mobilePhone = phone.number;
          else if (phone.type === "home")
            contact.homePhones = [...(contact.homePhones ?? []), phone.number];
          else if (phone.type === "business")
            contact.businessPhones = [
              ...(contact.businessPhones ?? []),
              phone.number,
            ];
        }
      }

      return graph.request("POST", "/me/contacts", contact);
    },
  }),

  defineTool({
    name: "update_contact",
    description: "Update contact properties.",
    inputSchema: {
      contact_id: z.string().describe("Contact ID to update"),
      updates: z
        .record(z.string(), z.unknown())
        .describe("Properties to update"),
    },
    confirmationPolicy: "always",
    handler: ({ graph }, params) => {
      validateId(params.contact_id, "contact_id");
      return graph.request(
        "PATCH",
        `/me/contacts/${params.contact_id}`,
        params.updates,
      );
    },
  }),

  defineTool({
    name: "delete_contact",
    description: "Delete a contact.",
    inputSchema: {
      contact_id: z.string().describe("Contact ID to delete"),
    },
    confirmationPolicy: "always",
    handler: async ({ graph }, params) => {
      validateId(params.contact_id, "contact_id");
      await graph.request("DELETE", `/me/contacts/${params.contact_id}`);
      return { success: true };
    },
  }),

  defineTool({
    name: "search_contacts",
    description: "Search contacts by name or email.",
    inputSchema: {
      query: z.string().describe("Search query string"),
      limit: z
        .number()
        .int()
        .min(1)
        .max(100)
        .optional()
        .describe("Max results. Defaults to 50."),
    },
    confirmationPolicy: "never",
    handler: async ({ graph }, params) => {
      const data = await graph.request("GET", "/me/contacts", undefined, {
        $search: `"${escapeKql(params.query)}"`,
        $top: String(params.limit ?? 50),
        $select:
          "id,displayName,givenName,surname,emailAddresses,businessPhones,mobilePhone",
      });
      return data.value ?? [];
    },
  }),
];
