import { NextResponse } from 'next/server';
import { ConfidentialClientApplication } from '@azure/msal-node';
import axios from 'axios';

const msalConfig = {
  auth: {
    clientId: process.env.OUTLOOK_CLIENT_ID,
    authority: `https://login.microsoftonline.com/${process.env.OUTLOOK_TENANT_ID}`,
    clientSecret: process.env.OUTLOOK_CLIENT_SECRET,
  },
};

const cca = new ConfidentialClientApplication(msalConfig);

export async function POST(request) {
  try {
    const bodyData = await request.json();
    const { name, email, message, phone, product_title, product_url, product_variant, seller_email } = bodyData;

    if (!name || !email || !message || !phone || !product_title || !product_url || !product_variant || !seller_email) {
      return NextResponse.json({ success: false, error: 'Missing required fields' }, { status: 400 });
    }

    const { accessToken } = await cca.acquireTokenByClientCredential({
      scopes: ['https://graph.microsoft.com/.default'],
    });

    const htmlContent = `
      <h2>New Product Enquiry for ${product_title}</h2>
      <p><strong>Name:</strong> ${name}</p>
      <p><strong>Email:</strong> ${email}</p>
      <p><strong>Phone:</strong> ${phone}</p>
      <div><strong>Product URL:</strong> <a href="${product_url}">${product_url}</a></div>

      <div><strong>Product Variant:</strong> ${product_variant}</div>
      <p><strong>Message:</strong></p>
      <p>${message}</p>
    `;

    await axios.post(
      `https://graph.microsoft.com/v1.0/users/${process.env.OUTLOOK_USER}/sendMail`,
      {
        message: {
          subject: `New Enquiry for ${product_title}`,
          body: {
            contentType: 'HTML',
            content: htmlContent,
          },
          toRecipients: [
            {
              emailAddress: {
                address: seller_email,
              },
            },
          ],
          replyTo: [
            {
              emailAddress: {
                address: email,
                name: name,
              },
            },
          ],
        },
        saveToSentItems: true,
      },
      {
        headers: {
          Authorization: `Bearer ${accessToken}`,
        },
      }
    );

    return NextResponse.json({ success: true, message: 'Email sent successfully!' }, { status: 200 });
  } catch (error) {
    console.error(error?.response?.data || error);
    return NextResponse.json({ success: false, error: 'Failed to send email' }, { status: 500 });
  }
}
 
