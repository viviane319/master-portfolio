# Experience Board: "Rooted in Community"

[`index.html`](index.html) is the live board (it scales to any screen); [`experience-board.png`](experience-board.png) is a 2200×1240 export to share or present.

## Swapping in real artifacts

Each colored card shows an image instead as soon as a file with its name exists in `images/`:

| File | Replaces the card for |
|------|-----------------------|
| `images/wsdp.jpg` | 88.1 WSDP "On Air" |
| `images/apia-vote.jpg` | APIA Vote Michigan |
| `images/light-house.jpg` | The Light House DaNang hostel |
| `images/gga-covid-report.jpg` | GG+A COVID-19 & Fundraising report |
| `images/cfjp.jpg` | Climate Funders Justice Pledge ($120M) |
| `images/nuf-georgetown.jpg` | National Urban Fellows / Georgetown |
| `images/do-better-comms.jpg` | Do Better Comms |
| `images/dear-crisis.jpg` | Dear Crisis, Sincerely… Grief |

After adding images, re-export the PNG:

```bash
"/Applications/Google Chrome.app/Contents/MacOS/Google Chrome" --headless=new --hide-scrollbars --window-size=2200,1240 --virtual-time-budget=5000 --screenshot=experience-board.png "file://$PWD/index.html?export"
```
