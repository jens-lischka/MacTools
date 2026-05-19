/* Office.js bridge — writes the generated newsletter into the message that
 * is currently being composed. `body.setAsync` replaces the entire body,
 * which is the right behaviour for a newsletter (it is the whole email). */

export function insertNewsletter(html: string): Promise<void> {
  return new Promise((resolve, reject) => {
    const item = Office.context.mailbox.item;
    if (!item) {
      reject(new Error("No message is open for composing."));
      return;
    }
    item.body.setAsync(
      html,
      { coercionType: Office.CoercionType.Html },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve();
        } else {
          reject(new Error(result.error?.message ?? "Could not update the email body."));
        }
      },
    );
  });
}
