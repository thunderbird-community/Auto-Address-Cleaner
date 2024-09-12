import ICAL from "/scripts/ical.js"
const debug = false;

browser.compose.onBeforeSend.addListener(async (tab, details) => {
    if (debug) {
        console.log(details);
        console.log("details.plainText of just after onBeforeSend: ", details.plainTextBody);
        console.log("original to:", details.to);
        console.log("original cc:", details.cc);
        console.log("original bcc:", details.bcc);
    }

    details.to = await stripDisplayName(details.to);
    details.cc = await stripDisplayName(details.cc);
    details.bcc = await stripDisplayName(details.bcc);

    if (debug) {
        console.log("striped");
        console.log("striped to:", details.to);
        console.log("striped cc:", details.cc);
        console.log("striped bcc:", details.bcc);
    }

    await browser.compose.setComposeDetails(tab.id, {
        to: details.to,
        cc: details.cc,
        bcc: details.bcc
    });
    let d = await browser.compose.getComposeDetails(tab.id);
    if (debug) {
        console.log("details of after setComposeDetails: ", d);
    };
});

async function expandListToEmails(name) {
    let books = await browser.addressBooks.list();
    for (let book of books) {
        let lists = await browser.addressBooks.mailingLists.list(book.id);
        let list = lists.find(l => l.name == name);
        if (list) {
            let contacts = await browser.addressBooks.mailingLists.listMembers(list.id);
            return contacts.map(c =>
                (new ICAL.Component(ICAL.parse(c.vCard))).getFirstPropertyValue("email")
            );
        }
    }
    return null;
}

async function stripDisplayName(addresses) {
    let emails = [];
    for (let address of addresses) {
        let parsed = await browser.messengerUtilities.parseMailboxString(address);
        // The parsed string could have multiple entries, each entry being a
        // ParsedMailbox (https://webextension-api.thunderbird.net/en/stable/messengerUtilities.html#parsedmailbox).
        for (let entry of parsed) {
            if (entry.email.includes("@")) {
                emails.push(entry.email);
                continue;
            }

            // The email did not include an @-tag. This may be a list.
            // Try to expand it and add the emails addr.
            let members = await expandListToEmails(entry.name);
            if (members) {
                emails.push(...members);
            } else {
                // Did not find the list, keep it as it is.
                emails.push(`${entry.name} <${entry.email}>`);
            }
        }
    }
    return emails.join(", ");
}
