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
    const appId = "ef1b8fd5-8e2d-4c65-9694-3d7c3e874b37";
    const clientSecret = "I4V8Q~X623nP6PrUE_Z_mZFFNGP2CVAkKGu8bbAO";
    const tenantId = "b0f8e84c-e4a8-4799-81b0-df150064037d";
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
