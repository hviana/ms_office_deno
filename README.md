# ms_office_deno

Microsoft Office Integration and Authentication for Deno.

## How to use

### Example:

You need to configure your server to receive authorization codes:

```typescript
import { req, res, Server } from "https://deno.land/x/faster/mod.ts";
import {
  MSOffice,
  MSOfficeApp,
} from "https://deno.land/x/ms_office_deno/mod.ts";
const server = new Server();
server.get(
  "/ms_auth_code/:customer",
  res("html"),
  async (ctx, next) => {
    if (ctx.params.customer && ctx.url.searchParams.get("code")) {
      await MSOffice.saveCode(
        ctx.params.customer, //The customer is a url parameter, ex (joe): /ms_auth_code/joe
        ctx.url.searchParams.get("code"), //The code is a url search parameter, ex: /ms_auth_code/joe?code=12372387238
      );
      await MSOffice.deleteUserToken(ctx.params.customer);
      ctx.res.body = `APP installed successfully.`;
    } else if (
      ctx.params.customer && ctx.url.searchParams.get("admin_consent") //url search parameter: /ms_auth_code/joe?admin_consent=True
    ) {
      const consent = ctx.url.searchParams.get("admin_consent") === "True";
      await MSOffice.saveAdminConsentStatus(ctx.params.customer, consent);
      if (consent) {
        await MSOffice.deleteUserToken(ctx.params.customer);
        ctx.res.body = `Successful admin consent.`;
      } else {
        ctx.res.body = "Failed to get admin consent.";
      }
    }
    await next();
  },
);
await server.listen({ port: 80 });
```

Now, you can use the Microsoft API:

```typescript
const customer = "joe_00203";
const officeApp = new MSOfficeApp(
  "YOUR Client Id",
  "YOUR Client Secret",
  `https://www.myexample.com/ms_auth_code/${customer}`, //put your server authorization code URL here
  customer,
  "YOUR App Tenant ID", //Tenant ID is optional with grant type => "authorization_code"
  "client_credentials", //grant type, values: "authorization_code" and "client_credentials", default is "client_credentials"
  ["https://graph.microsoft.com/.default"], //scope, default is: ["https://graph.microsoft.com/.default"]
  "https://graph.microsoft.com", //resource, default is: "https://graph.microsoft.com"
);
//NOW, YOU CAN USE:
//await officeApp.get(url), await officeApp.post(url, data), await officeApp.delete(url), etc.
```

You can generate the permission URLs with:

```
MSOffice.installUrl(host: string, redirect_path: string, customer: string, client_id: string, scope: string[])
```

```
MSOffice.adminConsentUrl(host: string, redirect_path: string, customer: string, client_id: string, scope: string[], tenantID: string)
```

## About

Author: Henrique Emanoel Viana, a Brazilian computer scientist, enthusiast of
web technologies, cel: +55 (41) 99999-4664. URL:
https://sites.google.com/view/henriqueviana

Improvements and suggestions are welcome!
