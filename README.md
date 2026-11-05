# Listening to Our Team

A Malaika Path tool for the **LEISH Directorate**: the Directorate of Labour, Employment Services, Industrial Relations and OSH at the Ministry of Gender, Labour and Social Development.

Staff sign up once with their name, email, role and unit, and receive a personal code by email to come back with. They answer six questions by speaking (English or Kiswahili) or typing. At the end they discover their reward: the first 3 full responses after launch get a Café Javas breakfast voucher (valid three months, any branch), and the next 30 a UGX 20,000 shopping voucher, collected in person from the Malaika Path project coordinator, Druscilla Kyeswa (0772 550 134), who is on site every week. Short answers go to the shopping voucher. During testing every tester receives the full reward email marked TEST, with a sample code; test conversations never take a real reward. Rewards are emailed from malaikapath@gmail.com (no reply). The Malaika Path team sees progress, gaps, quotes and recommendations in the control room.

| Who | Link |
|---|---|
| Staff | https://rebekahzadokmugisha-boop.github.io/lei-listening/ |
| Interviewer-led session | https://rebekahzadokmugisha-boop.github.io/lei-listening/?view=interview |
| Malaika Path control room | https://rebekahzadokmugisha-boop.github.io/lei-listening/?view=admin |

## How it is built

- `index.html` is the page (GitHub Pages).
- `config.js` holds the link to the backend.
- `backend/Code.gs` is the backend. It is a Google Apps Script owned by **malaikapath@gmail.com**, and it stores everything in the Google Sheet *Malaika Path · Listening Tool · Responses* (tabs: responses, settings, vouchers, staff, audit). Reward emails go out from that account, marked "no reply". They point staff to the project coordinator on 0772 550 134.

## Connecting the backend (one time, as malaikapath@gmail.com)

1. Go to script.google.com, open **New project**, and paste in `backend/Code.gs`.
2. Run `setup` and allow the permissions. The log shows the sheet link and the control-room password, which is stored as script property `ADMIN_KEY`.
3. Go to **Deploy > New deployment > Web app**. Set **Execute as** to Me and **Who has access** to Anyone, then copy the `/exec` link.
4. Paste that link into `config.js` as `apiUrl`.

## Loading voucher codes

In the control room, go to **Rewards > Load voucher codes**, choose the reward, and paste the codes one per line (a single column copied from Excel works). Each code is issued once, in completion order. Load each reward type separately. Only the Café Javas vouchers need codes; the shopping voucher gets its own collection code automatically.

## Day to day

- Everything is edited in the control room under **Settings**: questions, units, rewards, names, logos and the contact number.
- Voucher codes go under **Rewards**. People who are already waiting receive their code by email straight away.
- When the Directorate shares its staff list, paste it into the sheet's `staff` tab. To restrict the tool to listed staff, set **Who can take part** in Settings.
