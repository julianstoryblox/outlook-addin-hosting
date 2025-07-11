// functions.js
Office.onReady(info => {
  if (info.host === Office.HostType.Outlook) {
    // inizializzazioni eventuali
  }
});

function sendToN8N(event) {
  const item = Office.context.mailbox.item;
  item.body.getAsync(Office.CoercionType.Text, result => {
    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      console.error("Errore nel body:", result.error);
      event.completed();
      return;
    }
    const payload = {
      subject: item.subject,
      body: result.value,
      from: item.from?.emailAddress || "",
      to: item.to?.map(r => r.emailAddress).join(", "),
      cc: item.cc?.map(r => r.emailAddress).join(", ")
    };
    fetch("https://TUO-DOMINIO.COM/WEBHOOK-N8N", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload)
    })
    .then(r => {
      if (!r.ok) throw new Error(r.statusText);
      return Office.context.mailbox.item.notificationMessages.addAsync("ok", {
        type: "informationalMessage",
        message: "Email inviata a n8n!",
        icon: "icon16",
        persistent: false
      });
    })
    .catch(e => {
      console.error(e);
      Office.context.mailbox.item.notificationMessages.addAsync("err", {
        type: "errorMessage",
        message: "Errore invio a n8n",
        icon: "icon16",
        persistent: false
      });
    })
    .finally(() => event.completed());
  });
}
