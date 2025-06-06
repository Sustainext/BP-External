const fetch = require("node-fetch");
const { ConfidentialClientApplication } = require("@azure/msal-node");

exports.handler = async (event) => {
    if (event.httpMethod !== "POST") {
        return { statusCode: 405, body: "Only POST requests are allowed" };
    }

    let { name, phone, email_id } = {};
    try {
        ({ name, phone, email_id } = JSON.parse(event.body));
        if (!name || !phone || !email_id) throw new Error();
    } catch {
        return {
            statusCode: 400,
            body: "Invalid or missing fields in request body",
        };
    }

    // === Microsoft credentials ===
    const appId = process.env.APP_ID;
    const clientSecret = process.env.CLIENT_SECRET;
    const tenantId = process.env.TENANT_ID;
    const sender = "hello@sustainext.ai";

    const cca = new ConfidentialClientApplication({
        auth: {
            clientId: appId,
            authority: `https://login.microsoftonline.com/${tenantId}`,
            clientSecret,
        },
    });

    try {
        const { accessToken } = await cca.acquireTokenByClientCredential({
            scopes: ["https://graph.microsoft.com/.default"],
        });

        const htmlContent = `
      <html>
        <body>
          <h1>Hello ${name.split(" ")[0]},</h1>
          <p>Thanks for reaching out!</p>
          <p><strong>Name:</strong> ${name}</p>
          <p><strong>Email:</strong> ${email_id}</p>
          <p><strong>Phone:</strong> ${phone}</p>
        </body>
      </html>
    `;

        const mailData = {
            message: {
                subject: "Thanks for Reaching Out – Let's Catch Up Soon!",
                body: {
                    contentType: "HTML",
                    content: htmlContent,
                },
                toRecipients: [{ emailAddress: { address: email_id } }],
            },
            saveToSentItems: true,
        };

        const response = await fetch(
            `https://graph.microsoft.com/v1.0/users/${sender}/sendMail`,
            {
                method: "POST",
                headers: {
                    Authorization: `Bearer ${accessToken}`,
                    "Content-Type": "application/json",
                },
                body: JSON.stringify(mailData),
            }
        );

        if (!response.ok) {
            const errText = await response.text();
            return {
                statusCode: response.status,
                body: `Error sending email: ${errText}`,
            };
        }

        return { statusCode: 200, body: "Email sent successfully" };
    } catch (err) {
        return { statusCode: 500, body: `Internal error: ${err.message}` };
    }
};
