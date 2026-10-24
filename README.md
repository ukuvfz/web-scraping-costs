# how to extract data from website to excel automatically: 3 methods compared, plus the real cost of doing it at scale

The answer depends on one question: is the page you want built from plain HTML, or does it load its data with JavaScript? For a static Wikipedia table, Excel's built-in Power Query gets you there in two minutes and costs nothing. For Amazon, Google, or anything behind bot protection, Power Query sees an empty page — you need either a no-code scraper or a script that calls a scraping API, and the API route is the only one that keeps working after the site redesigns.

Here's the short version before we get into the details:

> **Just need one table from a static page?** Use Excel's Data → From Web or Google Sheets' IMPORTHTML. Free, minutes of setup.
>
> **Need recurring pulls from JS-heavy or protected sites?** A scraping API is the sturdier path. Most readers doing real volume land on ScraperAPI's Hobby plan at $49/month for 100,000 credits, after testing with the free 7-day trial. [👉 Start the free 5,000-credit trial (no credit card)](https://www.scraperapi.com/?fp_ref=coupons)

## Method 1: Excel and Google Sheets built-ins — free, but only for static pages

Excel's Power Query and Google Sheets' import formulas work only when the data sits in the raw HTML the server sends on first load. If your target is a plain HTML table — a league standings page, a government stats table, a Wikipedia list — start here.

In Excel (2016 and later), the path is **Data → From Web**. Paste the URL, pick the table in the Navigator window, and hit Load. The result is a live query: right-click to refresh, or set **Data → Queries & Connections → Properties → Refresh every N minutes** for scheduled pulls while the workbook stays open.

In Google Sheets, three formulas cover most cases:

- `=IMPORTHTML("https://en.wikipedia.org/wiki/List_of_countries_by_population_(United_Nations)", "table", 1)` — grabs the first table on a page
- `=IMPORTXML("https://example.com/product", "//h1")` — pulls a single element by XPath
- `=IMPORTXML(url, "//a/@href")` — collects every link on a page

Sheets re-evaluates these roughly once an hour. Fine for a daily snapshot, useless for anything resembling real time.

Three hard limits, none of which most tutorials spell out:

1. **No JavaScript rendering.** Power Query and IMPORTXML read only the initial HTML. On Amazon, Google result pages, and most e-commerce dashboards, what Excel sees is a blank shell. You get an empty table or a `#N/A` error.
2. **Shared IP blocking.** Google Sheets queries come from shared infrastructure that many sites block by default.
3. **Silent breakage.** These formulas point at a page's position in the DOM. When the site redesigns, they fail quietly — your spreadsheet just stops updating, and you find out when a report looks wrong.

## Method 2: No-code scrapers — good middle ground, per-site setup

Tools like Octoparse, Browse AI, and ParseHub handle JavaScript-rendered pages and pagination that Power Query can't touch. They differ in workflow, and that difference matters more than the feature lists suggest.

|  | How it works | Scheduling | Export | Watch-outs |
| --- | --- | --- | --- | --- |
| Octoparse | Desktop app + cloud tier, template or point-and-click workflow | Cloud runs on paid plans | CSV, Excel, databases, API | Each site needs its own task config |
| Browse AI | You train a "robot" by clicking fields; it re-runs on a schedule | Hourly/daily re-runs, syncs to Google Sheets or Airtable, change monitoring built in | CSV, JSON, Sheets, Airtable, S3 | Premium sites carry minimum credit requirements |
| ParseHub | Desktop click-based tool with a built-in browser for JS pages | Via paid plans | CSV, JSON, Google Sheets (needs a script) | Complex nested pagination is its strength |

The trade-off is maintenance. Every site gets its own configuration — ten sites means ten setups — and when a site changes its layout, the robot starts returning wrong data. You often discover this because the spreadsheet is already broken. Proxy rotation, where offered, sits behind paid tiers.

For a handful of sites on a predictable schedule, this works. Beyond that, the per-site config burden grows faster than most people expect.

## Method 3: Python plus a scraping API — the durable option

This is the route where a script fetches data through an API, the API handles the ugly parts (proxies, retries, CAPTCHAs, headless browsers), and pandas writes the result straight into a `.xlsx` file. Setup takes maybe 15 minutes once; after that, any target is just a different endpoint.

The core flow is three steps:

1. `pip install requests pandas openpyxl`
2. Send your target URL plus API key to the scraping service; get back structured JSON instead of raw HTML
3. Convert the JSON to a DataFrame and call `df.to_excel("output.xlsx", index=False)` — done, a finished Excel file

The reason this survives site redesigns is worth stating plainly: the API returns normalized fields (title, price, rating, link), so the page's internal structure changing doesn't break your script. With Power Query or a no-code robot, it does.

ScraperAPI is built for exactly this pattern. You send a request with your API key and the target URL; it rotates through an IP pool advertised at more than 40 million IPs across residential, datacenter, and mobile networks in 50+ countries, handles CAPTCHAs, and optionally renders JavaScript with `render=true`. Parameters like `country_code` geotarget requests, and `session_number` keeps the same IP across a sequence. For the big targets everyone actually scrapes, there are structured endpoints for Amazon product and search pages, Google Search and Shopping, and Walmart that return parsed JSON directly — you skip writing a parser entirely, which is what usually breaks when a site redesigns.

If you want to see how it behaves against your specific targets before spending anything, the free trial exists for that: [👉 Sign up for ScraperAPI's 7-day free trial with 5,000 credits](https://www.scraperapi.com/?fp_ref=coupons)

## ScraperAPI's full plan list and prices

ScraperAPI prices everything in credits. Here's every plan currently shown on the official pricing page:

| Plan | Monthly price | Annual (per month) | API credits/mo | Concurrent threads | Geo targeting | Buy |
| --- | --- | --- | --- | --- | --- | --- |
| Free | $0 | — | 1,000 | 5 | — | [ Get the free plan](https://www.scraperapi.com/?fp_ref=coupons) |
| Hobby | $49 | $44.10 | 100,000 | 20 | US & EU only | [ Start on Hobby](https://www.scraperapi.com/pricing?fp_ref=coupons) |
| Startup | $149 | $134.10 | 1,000,000 | 50 | US & EU only | [ View Startup plan details](https://www.scraperapi.com/pricing?fp_ref=coupons) |
| Business | $299 | $269.10 | 3,000,000 | 100 | Country-level | [ Check Business pricing](https://www.scraperapi.com/pricing?fp_ref=coupons) |
| Scaling | $475 | $427.50 | 5,000,000 | 200 | Country-level | [ See Scaling plan pricing](https://www.scraperapi.com/pricing?fp_ref=coupons) |
| Professional | $975 | $877.50 | 10,500,000 | 300 | Country-level | [ View Professional plan](https://www.scraperapi.com/pricing?fp_ref=coupons) |
| Advanced | $1,975 | $1,777.50 | 21,500,000 | 500 | Country-level | [ Check Advanced plan pricing](https://www.scraperapi.com/pricing?fp_ref=coupons) |
| Enterprise | Custom | Custom | 22,000,000+ | 500+ | Country-level | [ Talk to sales about Enterprise](https://www.scraperapi.com/?fp_ref=coupons) |

Annual billing takes 10% off the monthly rate. There's also a 7-day free trial with 5,000 credits, no credit card required, on top of the permanent free plan.

Every paid plan includes the same core feature set: JavaScript rendering, premium residential and mobile IPs, CAPTCHA and anti-bot handling, automatic retries, JSON auto-parsing, custom headers and sessions, unlimited bandwidth, and a 99.9% uptime guarantee. The plans differ in volume, concurrency, and how precisely you can target locations — Hobby and Startup only cover US and EU, so if your targets need other countries, Business is the actual entry point.

You can cancel anytime from the dashboard without being charged for cancelling.

## The part pricing pages don't highlight: what a credit actually buys

This is where budgets go sideways, so let's do the math the marketing page leaves out.

One credit equals one standard page request. But the sites people most often scrape are priced higher, and options stack:

| Request type | Credits |
| --- | --- |
| Standard page | 1 |
| Amazon | 5 |
| Google / Bing (incl. subdomains) | 25 |
| LinkedIn | 30 |
| JS rendering (`render=true`) | roughly +10 |
| Premium / ultra-premium IP pool | 10–30 |
| Bypassing Cloudflare, Datadome, PerimeterX | +10 |

So a protected-target request with rendering and a premium pool can eat 25–40 credits. Run that math on the Startup plan's 1,000,000 credits against something like Amazon search pages with rendering on, and you're covering somewhere in the range of 40,000–100,000 real pages — a fraction of what the credit count implies.

Two tools keep this under control, and both live in the dashboard:

- **Domain Cost Estimator** — check the exact per-request cost of any URL before you commit a budget to it
- **`max_cost`** — cap what any single request can spend, so a surprise parameter stack can't torch your balance

> The single most useful habit before choosing a plan: run your actual target URLs through the cost estimator, then divide your monthly page volume by the real per-page cost. Budget against that number, not against the credits on the pricing table.

When you've got that number, picking a tier is mechanical: [👉 Compare all ScraperAPI plans on the pricing page](https://www.scraperapi.com/pricing?fp_ref=coupons)

## What happens when you run out of credits

Default behavior: scraping pauses until your credits renew. That's actually the safe option for a scheduled Excel pipeline — no surprise charges, just a gap in your data.

If a gap isn't acceptable, there are two opt-in paths, depending on your plan:

- **Auto-upgrade** (Hobby through Professional): when credits run dry, your account moves to the next plan automatically. Turn it on deliberately; it doesn't happen by default.
- **Pay-as-you-go** (Scaling and up): keep scraping past your limit at a fixed per-credit rate, up to a monthly spending cap you set.

One more operational detail: exceeding your plan's concurrent thread limit returns a 429 and the request is rejected rather than queued. Nothing gets charged, but a job built for 50 concurrent requests will drop rows if it's running on Hobby's 20-thread limit. Match concurrency to the job, not the other way around.

## Who should skip this entirely

ScraperAPI is a developer-facing API. There's no point-and-click dashboard for selecting fields off a page — if that's what you need, a no-code tool like Octoparse or Browse AI fits better, and a free browser extension handles one-off static tables.

Even within the API use case, some limits are worth knowing upfront:

- Sticky sessions via `session_number` last roughly an hour. Account-based scraping or ad verification that needs a fixed IP for days won't work here.
- You can't pin specific IPs or ASNs. If a particular IP range gets blocked on your target, you're rerouting through the pool, not dialing in a specific address.
- Hobby and Startup get email/ticket support only — no live chat or phone.

The clear fit, based on what independent reviews consistently land on: a solo developer or small team running a price monitor, SERP tracker, or lead-gen scraper against standard e-commerce and search targets, at roughly 20,000–200,000 pages a month, who doesn't want to build and maintain proxy rotation logic themselves. If your targets aren't behind heavy bot protection, the default pool clears most requests without touching the expensive premium tiers at all — start there and only pay the multiplier when a target actually demands it. [👉 Test your targets on the free trial before picking a paid plan](https://www.scraperapi.com/?fp_ref=coupons)

## FAQ

**Can Excel pull data from a website automatically without any code?**
Yes, for static pages. Power Query (Data → From Web) creates a refreshable query, and you can schedule refreshes in the query properties. It fails on JavaScript-rendered and bot-protected sites.

**How much does it cost to extract 100,000 pages per month?**
With ScraperAPI, it depends on the target. Standard pages cost 1 credit each, so Hobby at $49/month covers 100,000 standard pages. The same volume on Google/Bing costs 25 credits per page — that's 2.5M credits, which puts you at Business ($299/month). Run your actual URLs through the Domain Cost Estimator to get the real number.

**Does the data land directly in Excel?**
Not by itself — the API returns JSON. Your script converts it with pandas and writes it out via `to_excel()`. Once written, that step never needs touching again, even when target sites redesign, because the JSON fields stay normalized.

**What's the difference between a scraping API and a browser extension scraper?**
Extensions operate on the page currently open in your browser — lowest friction, weakest at scale and scheduling. APIs handle volume, concurrency, proxies, and unattended runs, but require code. If your job is "this one table, today," an extension is faster; if it's "10,000 URLs without getting blocked," you want the API.

## Picking your path

If your data lives on static pages and refreshes while you're at your desk, Power Query is genuinely free and genuinely enough — no reason to pay for capability you won't use.

If you're watching prices on Amazon, tracking search rankings, or feeding an Excel report from sites that actively block scrapers, set up the Python-plus-API pipeline once and let it run. Start with the free trial, measure your real per-page cost with the estimator, then pick the smallest plan that covers your actual volume.

For most readers who've read this far, that's the Hobby plan at $49/month, with annual billing knocking it to $44.10 — and [👉 the 7-day free trial with 5,000 credits](https://www.scraperapi.com/?fp_ref=coupons) is the cheapest way to confirm that before the first invoice arrives.
