import { CFImap } from "cf-imap";

interface Env {
  KV: KVNamespace;
  OUTLOOK_USERNAME: string;
  OUTLOOK_PASSWORD: string;
}

export default {
  async scheduled(event: ScheduledEvent, env: Env, ctx: ExecutionContext) {
    const { OUTLOOK_USERNAME, OUTLOOK_PASSWORD } = env;
    const imap = new CFImap({
      host: "outlook.office365.com",
      port: 993,
      tls: true,
      auth: {
        username: OUTLOOK_USERNAME,
        password: OUTLOOK_PASSWORD,
      },
    });

    // https://docs.exerra.xyz/docs/npm-packages/cf-imap/v0.x.x/extendability
    try {
      await imap.connect();
      // const a = await imap.getFolders("", "*");
      // console.log(a);
      const selected = await imap.selectFolder("Inbox");
      const inboxSize = selected.emails as number;
      const in_emails = await imap.fetchEmails({
        limit: [inboxSize - 4, inboxSize],
        folder: "Inbox",
        fetchBody: false,
      });
      // console.log(emails);
      const in_info = in_emails.map(({ date, from }) => {
        return {
          date,
          from: from.match(/(?<=<).+?(?=>)/)![0]
        };
      });
      console.log("Latest Inbox:", in_info);

      const sent = await imap.selectFolder("Sent");
      const sentSize = sent.emails as number;
      const sent_emails = await imap.fetchEmails({
        limit: [sentSize - 4, sentSize],
        folder: "Sent",
        fetchBody: false,
      });
      const sent_info = sent_emails.map(({ date, to }) => {
        return {
          date,
          to: to.match(/(?<=<).+?(?=>)/)![0]
        };
      });
      console.log("Latest Sent:", sent_info);

      // await imap.logout();
      // It's buggy :(
      const query = `A023 LOGOUT\r\n`;
      const encoded = imap.encoder.encode(query);
      await imap.writer!.write(encoded);
      const { value } = await imap.reader!.read();
      console.log(imap.decoder.decode(value));

      await env.KV.put("box", JSON.stringify({ in_info, sent_info }));
    }
    // Don't await! Why is that?
    // https://developers.cloudflare.com/workers/runtime-apis/tcp-sockets/#close-tcp-connections
    finally { imap.socket!.close(); }
  }
};
