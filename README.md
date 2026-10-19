# ED Shirt Order

Shirt order form for the Emergency Department team. The form is a static page on GitHub Pages. Orders go to a private Google Sheet through a Google Apps Script web app, which also sends each person a confirmation email.

**No order data lives in this repo.** Names, emails and sizes are stored only in the Google Sheet, which only you (and anyone you share it with) can open. The web app accepts new orders but never sends any order data back out.

## Files

| File | What it does |
| --- | --- |
| `index.html` | The order form |
| `config.js` | The web app URL and team code setting |
| `apps-script/Code.gs` | Backend: saves orders to the sheet and sends confirmation emails |

## Setup

1. **Create the sheet.** In Google Drive, create a new Google Sheet named `ED Shirt Orders`. Keep it private.
2. **Add the script.** In the sheet, open **Extensions → Apps Script**. Delete the sample code, paste in all of `apps-script/Code.gs`, and save.
3. **Build the tabs.** Pick `setup` from the function menu and click **Run**. Approve the permissions when asked. This creates the **Orders** tab and a **Totals** tab that tallies shirts by size.
4. **(Optional) Add a team code.** Go to **Project Settings → Script Properties** and add `TEAM_CODE` with a code you'll share only with ED staff. Orders without the right code are rejected.
5. **Deploy.** Click **Deploy → New deployment**, choose the **Web app** type, and set:
   - Execute as: **Me**
   - Who has access: **Anyone**

   "Anyone" lets people submit the form without a Google sign-in. It does **not** give them access to the sheet. Copy the web app URL that ends in `/exec`.
6. **Connect the form.** In `config.js`, paste the URL into `endpoint`. If you set a team code, change `teamCodeRequired` to `true`. Commit and push.
7. **Turn on Pages.** In the repo, go to **Settings → Pages** and publish from the `main` branch, root folder.

## How orders work

- Each email address gets one order. Submitting again with the same email updates that order and sends an "updated" confirmation.
- The **Totals** tab shows how many of each sleeve style and size to order from the printer.
- Each order row has one column per style and size (SS = short sleeve, LS = long sleeve).
- Confirmation emails come from your Google account and show as "ED Shirt Order".

## Security notes

- The sheet is never public. Only people you share it with can see orders.
- The web app only writes orders and never returns anyone else's data.
- Inputs are length-limited and cleaned so nothing typed into the form can run as a spreadsheet formula.
- A hidden field filters out spam bots. Repeat submissions from the same email are rate-limited.
- The page includes `noindex` so search engines skip it.
- After changing `Code.gs`, use **Deploy → Manage deployments → Edit → New version** so the URL stays the same.

## Shirt photos

The form shows a placeholder for each style until photos are added. To add them, put the image files in an `images/` folder (for example `images/short-sleeve.jpg` and `images/long-sleeve.jpg`) and set their paths under `images` in `config.js`. Square images look best.
