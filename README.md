# NeighborTech SimpleQuote

**A free, open-source CRM that turns Google Sheets and Google Calendar into a practical customer, booking, quoting, and invoicing system for small businesses.**

> Project status: working MVP undergoing live Google Sheets and Calendar testing.

A free, Google Sheets-backed small-business workflow that combines:

**Lead → Booking → Quote → Invoice → Payment tracking**

## Why this project exists

NeighborTech SimpleQuote helps small businesses manage customer inquiries, appointments, quotes, and invoices without starting with an expensive CRM subscription. Each business controls its own Google Sheet and Google Calendar connection, while the project remains reusable and customizable.

Many independent businesses need more than a contact form but do not need the cost or complexity of a large CRM. NeighborTech SimpleQuote brings the everyday workflow together in one place: capture a website inquiry or phone call, create or match the customer, check real calendar availability, book an appointment, build a quote from saved products and services, convert an accepted quote into an invoice, and follow the customer’s history from first contact through payment.

The project is designed around ownership and accessibility. Business data stays in the business owner’s Google Sheet, appointments connect to their Google Calendar, and the interface can be adapted for different services, prices, terminology, branding, and workflows. The goal is to give local service providers, consultants, freelancers, tradespeople, and community businesses a useful starting point they can understand, operate, and extend without being locked into a monthly software subscription.

## Search keywords

`free CRM` · `open-source CRM` · `Google Sheets CRM` · `small-business CRM` · `booking system` · `appointment scheduling` · `Google Calendar booking` · `quote software` · `estimate software` · `invoice software` · `customer management` · `lead management` · `service-business software` · `Google Apps Script` · `JavaScript CRM`

This MVP combines and extends three Neighbor Tech projects:

- [lead-capture](https://github.com/christianpgonz-tech/lead-capture)
- [booking-request](https://github.com/christianpgonz-tech/booking-request)
- SimpleQuote quote/invoice prototype

## MVP features

- Customizable public lead form
- Shared customer directory with automatic email/phone matching
- Fast manual customer and phone-inquiry entry
- Customer history across leads, bookings, quotes, and invoices
- Real Google Calendar availability and booking
- Monthly booking calendar with previous/next navigation
- Booking filters for service and status, plus calendar/list views
- Book from the CRM while speaking with a customer, with a final conflict check
- Unified dashboard tabs for leads, bookings, quote requests, quotes, and invoices
- Products & Services catalog with price, unit, duration, description, type, and tax setting
- One-click catalog items in quote line items
- Create a pre-filled quote from a lead or booking
- Create quotes and invoices directly
- Convert accepted quotes into independent invoices
- Track paid amount, method, and balance due
- Branded printable quote/invoice view
- Business settings wizard
- Installable PWA shell
- One isolated Google Sheet per business
- Smart Notes Assistant with editable extraction of customer, service, preferred date, and budget
- Browser-based Hugging Face service matching with no paid API or exposed token
- One-click creation of a customer, lead, booking, or draft quote from reviewed notes

## Try the demo before connecting Google

The dashboard includes a browser-only demo mode with sample data. Serve the `website` folder from any local static server and open:

```text
app.html?demo=1
```

Demo mode lets you explore leads, bookings, quote requests, quotes, invoices, pre-filled quote creation, and the Smart Notes Assistant. Nothing is sent to Google and changes reset when the page refreshes.

Opening `app.html` without `?demo=1` starts connected mode and requires the Apps Script settings described below.

## Architecture

```text
Public lead form ───────┐
Public booking form ────┼──> Google Apps Script ──> Google Sheet
Public quote request ───┘            │                   │
                                     └──> Google Calendar │
                                                          │
Private dashboard <────────────────────────────────────────┘
```

The Google Sheet is the pilot database. The optional Smart Notes Assistant runs a Hugging Face model in the visitor's browser to match notes with the service catalog. Contact details, dates, and budgets are extracted locally, all results remain editable, and a built-in fallback keeps the workflow usable if the model cannot load.

## Install

### 1. Create the spreadsheet

Create a blank Google Sheet. Import every CSV in `sheet-templates/` as a separate tab. The CSV filenames now match the intended tab names exactly:

- `Leads`
- `Customers`
- `Bookings`
- `Products_Services`
- `Quote_Requests`
- `Quotes`
- `Quote_Items`
- `Invoices`
- `Invoice_Items`
- `Settings`

The backend also tolerates capitalization and separator differences. For example, `Quote_Items`, `Quote-Items`, and `quote items` resolve to the same tab. The underscore names above remain the recommended convention.

### 2. Install Apps Script

Open **Extensions → Apps Script**, replace the starter code with `apps-script/Code.gs`, and save.

In **Project Settings → Script properties**, add:

```text
ADMIN_KEY = a long private random value
```

Configure `SERVICES`, `CALENDAR_ID`, and `BOOKING_WINDOW_DAYS` near the top of `Code.gs`, or use the included booking builder and copy its service configuration.

Deploy it as a web app:

- Execute as: **Me**
- Access: **Anyone** (required for customer-facing forms)

Keep the `/exec` URL.

### 3. Configure the website files

Add the `/exec` URL to:

- `website/app.html`
- `website/settings-wizard.html`
- `website/forms/lead-form.html`
- `website/forms/booking-form.html`
- the generated quote-request form from `website/request-builder.html`

Add the same private `ADMIN_KEY` only to:

- `website/app.html`
- `website/settings-wizard.html`

Do not add the admin key to public forms.

### 4. Host carefully

The public forms may be hosted on a normal static website. Treat `app.html`, `settings-wizard.html`, and `print-view.html` as private owner tools. Do not publish a configured owner dashboard containing the admin key in a public GitHub repository or public site.

For this MVP, the practical approach is:

1. keep placeholders in the public GitHub repository;
2. make a private configured deployment copy for the business owner;
3. publish only the customer forms publicly.

GitHub Pages is suitable for documentation and a sanitized demo, not for running a commercial SaaS or handling sensitive transactions.

## Repository map

```text
apps-script/Code.gs             unified backend
website/app.html                owner dashboard
website/settings-wizard.html    branding and document settings
website/print-view.html         printable documents
website/request-builder.html    quote-request form builder
website/forms/                  public lead and booking forms
website/builders/               visual lead and booking builders
sheet-templates/                Google Sheet tab templates
docs/                           scope, security notes, and roadmap
```

## Security boundary

The public surface is limited to:

- lead submission;
- booking service list and availability;
- booking submission;
- quote-request submission.

Dashboard reads and writes require the Apps Script `ADMIN_KEY`. This is an MVP protection layer, not full multi-user authentication. Before offering a centrally hosted product to unrelated businesses, move the owner interface behind managed authentication and migrate to a proper multi-tenant database.

## Known MVP limitations

- Admin authorization is one shared key, not named user accounts.
- Quote/invoice dashboard requests still use JSONP because of Apps Script browser restrictions.
- Quote requests still use the original JSONP submission flow.
- Customer records are embedded in source records and documents; a dedicated Customers module is next.
- Payment tracking uses an aggregate paid amount rather than a payment ledger.
- Expenses, mileage, reporting, and additional AI workflows are roadmap items.
- Existing lead and booking rows must receive unique IDs before they can be converted from the dashboard.

## Migrating existing lead and booking rows

The new templates add `Lead_ID` and `Booking_ID`. Add those columns to existing tabs. Give every existing row a unique value such as `LEAD-0001` or `BKG-0001`. New submissions receive IDs automatically.

## Next milestone

1. Payments ledger
2. Expenses and mileage
3. Dashboard summary cards
4. Full setup wizard for terminology and modules
5. Hugging Face “rough notes → proposed records” assistant with review-before-save
