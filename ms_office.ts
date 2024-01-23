/*
Created by: Henrique Emanoel Viana
Githu: https://github.com/hviana
Page: https://sites.google.com/view/henriqueviana
cel: +55 (41) 99999-4664
*/
export class MSOffice {
  static ms_timeout = 300; //if there are xxx seconds left before the token expires, renew
  static kvDatabase: Deno.Kv = undefined;

  static async initDB() {
    MSOffice.kvDatabase = await Deno.openKv();
  }
  static async installUrl(
    host: string,
    redirect_path: string,
    customer: string,
    client_id: string,
    scope: string[],
  ) {
    return `https://login.windows.net/common/oauth2/authorize?response_type=code&response_mode=query&scope=${
      encodeURIComponent(scope.join(" "))
    }&client_id=${client_id}&redirect_uri=${
      encodeURIComponent(`${host}/${redirect_path}/${customer}`)
    }`;
  }
  static async adminConsentUrl(
    host: string,
    redirect_path: string,
    customer: string,
    client_id: string,
    scope: string[],
    tenantID: string,
  ) {
    return `https://login.microsoftonline.com/${tenantID}/v2.0/adminconsent?response_type=code&response_mode=query&scope=${
      encodeURIComponent(scope.join(" "))
    }&client_id=${client_id}&redirect_uri=${
      encodeURIComponent(`${host}/${redirect_path}/${customer}`)
    }`;
  }

  static async saveCode(customer: string, code: string) {
    await MSOffice.initDB();
    await MSOffice.kvDatabase.set(["ms_auth_code", customer], code);
  }
  static async saveAdminConsentStatus(customer: string, consent: boolean) {
    await MSOffice.initDB();
    await MSOffice.kvDatabase.set(["ms_admin_consent", customer], consent);
  }
  static async deleteUserToken(customer: string) {
    await MSOffice.initDB();
    await MSOffice.kvDatabase.delete(["ms_token", customer]);
  }

  static async getToken(
    client_id: string,
    client_secret: string,
    redirect_uri: string, //sometimes needs "/" in the end
    customer: string,
    grant_type: string,
    scope: string[],
    resource: string,
    tenantID: string,
  ): Promise<string> {
    await MSOffice.initDB();
    var refresh_token = "";
    const userRef = grant_type === "authorization_code"
      ? "." + await MSOffice.kvDatabase.get(["ms_auth_code", customer])
      : "";
    const tokenData =
      await MSOffice.kvDatabase.get(["ms_token", customer, userRef]) || {};
    if (tokenData.refresh_token) {
      if (
        ((Math.round(Date.now() / 1000) + MSOffice.ms_timeout) -
          tokenData.time) >=
          tokenData.expires_in
      ) {
        refresh_token = tokenData.refresh_token;
      } else {
        return tokenData.access_token;
      }
    }
    const data: any = {
      client_id: client_id,
      scope: scope.join(" "),
      redirect_uri: redirect_uri,
      client_secret: client_secret,
    };
    if (refresh_token) {
      data["refresh_token"] = refresh_token;
      data["grant_type"] = "refresh_token";
    } else {
      if (grant_type === "authorization_code") {
        data["code"] = await MSOffice.kvDatabase.get([
          "ms_auth_code",
          customer,
        ]);
        data["resource"] = resource;
      }
      data["grant_type"] = grant_type;
    }
    const result = await fetch(
      grant_type !== "authorization_code"
        ? `https://login.microsoftonline.com/${tenantID}/oauth2/v2.0/token`
        : "https://login.windows.net/common/oauth2/token",
      {
        method: "POST",
        headers: {
          "Content-Type": "application/x-www-form-urlencoded",
          "Accept": "application/json",
        },
        body: (new URLSearchParams(data)).toString().replace(/\+/g, "%20"),
      },
    );
    const res = await result.json();
    if (res.access_token) {
      await MSOffice.kvDatabase.set(["ms_token", customer, userRef], {
        ...res,
        time: Math.round(Date.now() / 1000),
      });
    } else if (Object.keys(tokenData).length > 0) {
      await MSOffice.kvDatabase.delete(["ms_token", customer, userRef]);
      return await MSOffice.getToken(
        client_id,
        client_secret,
        redirect_uri,
        customer,
        grant_type,
        scope,
        resource,
        tenantID,
      );
    }
    if (!res.access_token) {
      throw new Error(`Error in get token, data: ${JSON.stringify(res)}`);
    }
    return res.access_token;
  }
}
export class MSOfficeApp {
  client_id: string;
  client_secret: string;
  redirect_uri: string;
  customer: string;
  tenantID: string;
  grant_type: string;
  scope: string[];
  resource: string;
  constructor(
    client_id: string,
    client_secret: string,
    redirect_uri: string,
    customer: string,
    tenantID: string = "",
    grant_type: string = "client_credentials",
    scope: string[] = [
      "https://graph.microsoft.com/.default",
    ],
    resource: string = "https://graph.microsoft.com",
  ) {
    this.client_id = client_id;
    this.client_secret = client_secret;
    this.redirect_uri = redirect_uri;
    this.customer = customer;
    this.tenantID = tenantID;
    this.grant_type = grant_type;
    this.scope = scope;
    this.resource = resource;
  }
  async getToken() {
    return await MSOffice.getToken(
      this.client_id,
      this.client_secret,
      this.redirect_uri,
      this.customer,
      this.grant_type,
      this.scope,
      this.resource,
      this.tenantID,
    );
  }
  async get(endpoint: string) {
    return await this.request(endpoint, "GET", null);
  }
  async put(endpoint: string, data: any = {}) {
    return await this.request(endpoint, "PUT", data);
  }
  async post(endpoint: string, data: any = {}) {
    return await this.request(endpoint, "POST", data);
  }
  async delete(endpoint: string, data: any = {}) {
    return await this.request(endpoint, "DELETE", data);
  }
  async request(
    endpoint: string,
    method: string = "GET",
    data: any = {},
  ): Promise<any> {
    const headers = new Headers({
      "Content-Type": "application/json",
      "Accept": "application/json",
      "Authorization": `Bearer ${await this.getToken()}`,
    });
    const reqData = JSON.stringify(data);
    var params: any = {
      method: method,
      headers: headers,
    };
    if (method !== "GET" && method !== "HEAD") {
      params.body = reqData;
    }
    var request = await fetch(
      `${this.resource}/${endpoint}`,
      params,
    );
    var res: any = {};
    try {
      res = await request.json();
    } catch (e) {}
    if (res.error) {
      throw new Error(res.error.message);
    }
    if (res.ok !== undefined && !res.ok) {
      throw new Error(res.statusText);
    }
    return res;
  }
  async createChat(toList: string[]) {
    const data: any = {
      "chatType": toList.length > 2 ? "group" : "oneOnOne",
      "members": [],
    };
    for (const to of toList) {
      data["members"].push(
        {
          "@odata.type": "#microsoft.graph.aadUserConversationMember",
          "roles": ["owner"],
          "user@odata.bind": `https://graph.microsoft.com/v1.0/users('${to}')`,
        },
      );
    }
    return await this.post(
      "v1.0/chats",
      data,
    );
  }
  async sendToTeams(
    msg: string,
    teamId: string = "",
    channelId: string = "",
    chatId: string = "",
    to: string[] = [],
    files: { name: string; content: Uint8Array }[] = [],
  ) {
    if (!chatId && !teamId && !channelId) {
      chatId = (await this.createChat(to)).id;
    }
    const data: any = {
      body: {
        "contentType": "html",
        "content": msg,
      },
    };
    if (files.length > 0) {
      data["attachments"] = [];
    }
    for (const file of files) {
      data["attachments"].push({
        "@odata.type": "#microsoft.graph.fileAttachment",
        "name": file.name,
        "contentType": "application/octet-stream",
        "contentBytes": await this.uint8ArrayToBase64(file.content),
      });
    }
    return await this.post(
      chatId
        ? `v1.0/chats/${chatId}/messages`
        : `v1.0/teams/${teamId}/channels/${channelId}/messages`,
      data,
    );
  }
  async uint8ArrayToBase64(content: Uint8Array) {
    // use a FileReader to generate a base64 data URI:
    const base64url: any = await new Promise((r) => {
      const reader = new FileReader();
      reader.onload = () => r(reader.result);
      reader.readAsDataURL(new Blob([content]));
    });
    // remove the `data:...;base64,` part from the start
    return base64url.slice(base64url.indexOf(",") + 1);
  }
  async sendMail(
    from: string | null,
    toList: string[],
    subject: string,
    html: string = "",
    files: { name: string; content: Uint8Array }[] = [],
  ) {
    const data: any = {
      Message: {
        Subject: subject,
        Body: {
          "ContentType": "HTML",
          "Content": html,
        },
        ToRecipients: [],
      },
      SaveToSentItems: "false", //default true
    };
    if (files.length > 0) {
      data["Message"]["attachments"] = [];
    }
    for (const file of files) {
      data["Message"]["attachments"].push({
        "@odata.type": "#microsoft.graph.fileAttachment",
        "name": file.name,
        "contentType": "application/octet-stream",
        "contentBytes": await this.uint8ArrayToBase64(file.content),
      });
    }
    for (const to of toList) {
      data["Message"]["ToRecipients"].push({ EmailAddress: { Address: to } });
    }
    return await this.post(
      from ? `v1.0/users/${from}/sendMail` : "v1.0/me/sendMail",
      data,
    );
  }
}
