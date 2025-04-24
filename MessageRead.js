/* MessageRead.js – v39
   Changes in v36:
   • Arrows inverted in HTML/CSS (see .chevron transforms there).
   • Added "kaseya.net" to verifiedDomains set.

   Changes in v37:
   • Attachments separated into their own collapsible card (#attachments-card).
   • The click-handler for #attachBadgeContainer now expands #attachments-card instead of #threats-card.

   Changes in v38 (internal domain trust):
   • Added logic to trust internal (business) senders if envelope domain and From domain both match the user's
     own business domain, and if SPF/DKIM/DMARC are "pass" (non-personal domain).
   • Introduced window.__userDomain and window.__internalSenderTrusted to coordinate this logic in checkAuthHeaders
     and senderClassification without removing or breaking existing code.

   Changes in v39 (fix internal trust display):
   • The classification was happening before checkAuthHeaders finished. Now, after we set
     window.__internalSenderTrusted in the async headers callback, we re-run senderClassification(it)
     (and fromSenderMismatch(it) again) to immediately reflect the updated "Verified Sender" badge.
*/

(function () {
    "use strict";

    /* ---------- 1. CONSTANTS ---------- */
    const THEME_KEY = "bkEmailAddinTheme";
    const verifiedSenders = [
        "support@microsoft.com",
        "support@amazon.com",
        "support@google.com"
    ];

    // A large set of reputable-company domains for domain-based verification:
    const verifiedDomains = new Set([
        // 1. E-commerce Market Leaders (20)
        "amazon.com", "ebay.com", "alibaba.com", "aliexpress.com", "jd.com", "walmart.com", "target.com", "rakuten.com", "mercadolibre.com", "flipkart.com", "overstock.com", "etsy.com", "groupon.com", "wayfair.com", "zappos.com", "shein.com", "gearbest.com", "banggood.com", "tmall.com", "shopify.com",

        // 2. Large Retailers & Department Stores (20)
        "costco.com", "kohls.com", "bestbuy.com", "macys.com", "nordstrom.com", "bloomingdales.com", "dillards.com", "jcpenney.com", "sears.com", "neimanmarcus.com", "saksfifthavenue.com", "meijer.com", "biglots.com", "rossstores.com", "tjmaxx.com", "marshalls.com", "burlington.com", "dollargeneral.com", "familydollar.com", "bedbathandbeyond.com",

        // 3. Fashion & Apparel (20)
        "gap.com", "oldnavy.com", "bananarepublic.com", "uniqlo.com", "hm.com", "zara.com", "forever21.com", "asos.com", "revolve.com", "urbanoutfitters.com", "freepeople.com", "anthropologie.com", "abercrombie.com", "hollisterco.com", "fashionnova.com", "victoriassecret.com", "adidas.com", "nike.com", "underarmour.com", "lululemon.com",

        // 4. Technology & Software (20)
        "microsoft.com", "apple.com", "google.com", "oracle.com", "sap.com", "salesforce.com", "adobe.com", "ibm.com", "intel.com", "dell.com", "hp.com", "lenovo.com", "asus.com", "nvidia.com", "amd.com", "autodesk.com", "zoom.us", "slack.com", "gitlab.com", "atlassian.com",

        /* Inserted here: "kaseya.net" */
        "kaseya.net",

        // 5. Electronics & Hardware (20)
        "samsung.com", "lg.com", "sony.com", "panasonic.com", "philips.com", "sharpusa.com", "huawei.com", "xiaomi.com", "oneplus.com", "realme.com", "oppo.com", "vivo.com", "toshiba.com", "pioneer.com", "jvc.com", "canon.com", "nikon.com", "epson.com", "fujifilm.com", "bose.com",

        // 6. Payment & Financial Services (20)
        "paypal.com", "stripe.com", "squareup.com", "venmo.com", "skrill.com", "payoneer.com", "wepay.com", "adyen.com", "authorize.net", "alipay.com", "neteller.com", "googlepay.com", "amazonpay.com", "worldpay.com", "firstdata.com", "payu.com", "bill.com", "intuit.com", "xero.com", "coinbase.com",

        // 7. Banks & Lending (20)
        "chase.com", "wellsfargo.com", "bankofamerica.com", "citi.com", "usbank.com", "pnc.com", "truist.com", "capitalone.com", "americanexpress.com", "discover.com", "goldmansachs.com", "barclays.com", "hsbc.com", "lloydsbank.com", "rbs.co.uk", "santander.com", "bbva.com", "bnymellon.com", "sofi.com", "ally.com",

        // 8. Insurance (20)
        "geico.com", "progressive.com", "allstate.com", "statefarm.com", "farmers.com", "usaa.com", "libertymutual.com", "nationwide.com", "travelers.com", "chubb.com", "zurichna.com", "thehartford.com", "metlife.com", "prudential.com", "aetna.com", "cigna.com", "humana.com", "aflac.com", "coloniallife.com", "globelife.com",

        // 9. Healthcare & Pharma (20)
        "pfizer.com", "moderna.com", "johnsonandjohnson.com", "merck.com", "astrazeneca.com", "novartis.com", "roche.com", "gsk.com", "sanofi.com", "abbvie.com", "bristolmyerssquibb.com", "lilly.com", "bayer.com", "amgen.com", "teva.com", "viatris.com", "regeneron.com", "cardinalhealth.com", "mckesson.com", "abbott.com",

        // 10. Telecom & ISPs (20)
        "att.com", "verizon.com", "t-mobile.com", "sprint.com", "xfinity.com", "comcast.com", "charter.com", "spectrum.com", "centurylink.com", "frontier.com", "bt.com", "vodafone.com", "orange.com", "telefonica.com", "rogers.com", "bell.ca", "telus.com", "telstra.com", "mtn.com", "uscellular.com",

        // 11. Social Media & Networking (20)
        "facebook.com", "instagram.com", "twitter.com", "linkedin.com", "snapchat.com", "pinterest.com", "tiktok.com", "reddit.com", "tumblr.com", "weibo.com", "wechat.com", "discord.com", "quora.com", "meetup.com", "xing.com", "vk.com", "flickr.com", "behance.net", "deviantart.com", "medium.com",

        // 12. Internet & Tech Giants (20)
        "baidu.com", "yandex.com", "cloudflare.com", "akamai.com", "digitalocean.com", "rackspace.com", "godaddy.com", "namecheap.com", "wordpress.com", "squarespace.com", "weebly.com", "wix.com", "bigcommerce.com", "mailchimp.com", "hubspot.com", "constantcontact.com", "webex.com", "cisco.com", "github.com", "tencent.com",

        // 13. Travel Sites (20)
        "booking.com", "expedia.com", "tripadvisor.com", "orbitz.com", "travelocity.com", "priceline.com", "kayak.com", "skyscanner.com", "trivago.com", "hotwire.com", "hopper.com", "agoda.com", "cheapoair.com", "ebookers.com", "cheapair.com", "airfarewatchdog.com", "lastminute.com", "travelzoo.com", "travelgenio.com", "momondo.com",

        // 14. Airlines (20)
        "delta.com", "united.com", "southwest.com", "american.com", "aa.com", "alaskaair.com", "jetblue.com", "spirit.com", "hawaiianairlines.com", "allegiantair.com", "britishairways.com", "lufthansa.com", "airfrance.com", "klm.com", "emirates.com", "qatarairways.com", "etihad.com", "cathaypacific.com", "singaporeair.com", "aerlingus.com",

        // 15. Hotels & Accommodation (20)
        "marriott.com", "hilton.com", "hyatt.com", "ihg.com", "choicehotels.com", "wyndhamhotels.com", "accor.com", "ritzcarlton.com", "fourseasons.com", "fairmont.com", "starwoodhotels.com", "mgmresorts.com", "wynnresorts.com", "hostels.com", "motel6.com", "bestwestern.com", "radissonhotels.com", "scandichotels.com", "oyorooms.com", "airbnb.com",

        // 16. Car Rentals & Transportation (20)
        "hertz.com", "avis.com", "budget.com", "enterprise.com", "alamo.com", "nationalcar.com", "thrifty.com", "dollar.com", "sixt.com", "uhaul.com", "pensketruckrental.com", "lyft.com", "uber.com", "grab.com", "bolt.eu", "cabify.com", "lime.me", "bird.co", "spin.app", "turo.com",

        // 17. Food & Beverage (20)
        "starbucks.com", "dunkindonuts.com", "mcdonalds.com", "burgerking.com", "wendys.com", "tacobell.com", "pizzahut.com", "dominos.com", "papajohns.com", "chipotle.com", "panerabread.com", "chick-fil-a.com", "kfc.com", "subway.com", "fiveguys.com", "sonicdrivein.com", "arbys.com", "dairyqueen.com", "littlecaesars.com", "jimmyjohns.com",

        // 18. Logistics & Shipping (20)
        "ups.com", "fedex.com", "dhl.com", "usps.com", "canadapost.ca", "royalmail.com", "parcelforce.com", "hermesworld.com", "dpd.com", "tnt.com", "aramex.com", "gls-group.eu", "yamato-hd.co.jp", "japanpost.jp", "laposte.fr", "upsupplychain.com", "fedexcustomcritical.com", "dhlglobalforwarding.com", "ontrac.com", "yrc.com",

        // 19. Media & Entertainment (20)
        "netflix.com", "hulu.com", "disneyplus.com", "hbo.com", "showtime.com", "paramountplus.com", "peacocktv.com", "discoveryplus.com", "espn.com", "fox.com", "abc.com", "nbc.com", "cbs.com", "bbc.co.uk", "cnn.com", "bloomberg.com", "reuters.com", "theguardian.com", "nytimes.com", "wsj.com",

        // 20. Automotive (20)
        "ford.com", "gm.com", "chevrolet.com", "toyota.com", "honda.com", "nissanusa.com", "hyundaiusa.com", "kia.com", "tesla.com", "bmw.com", "mercedes-benz.com", "audi.com", "volkswagen.com", "porsche.com", "volvo.com", "subaru.com", "mazdausa.com", "dodge.com", "jeep.com", "ramtrucks.com",

        // 21. Education (20)
        "harvard.edu", "mit.edu", "stanford.edu", "berkeley.edu", "ox.ac.uk", "cam.ac.uk", "yale.edu", "princeton.edu", "columbia.edu", "ucla.edu", "nyu.edu", "upenn.edu", "caltech.edu", "cmu.edu", "gatech.edu", "uf.edu", "umich.edu", "k12.com", "coursera.org", "edx.org",

        // 22. Nonprofits & International Orgs (20)
        "un.org", "who.int", "worldbank.org", "imf.org", "wto.org", "unesco.org", "unicef.org", "redcross.org", "salvationarmy.org", "unitedway.org", "habitat.org", "wwf.org", "greenpeace.org", "amnesty.org", "doctorswithoutborders.org", "care.org", "oxfam.org", "mercycorps.org", "charitywater.org", "worldvision.org",

        // 23. Government & Public Services (20)
        "usa.gov", "irs.gov", "ssa.gov", "nps.gov", "nasa.gov", "gov.uk", "canada.ca", "australia.gov.au", "india.gov.in", "gov.cn", "europa.eu", "whitehouse.gov", "senate.gov", "house.gov", "justice.gov", "ny.gov", "ca.gov", "gov.za", "scot.gov", "uscis.gov",

        // 24. Manufacturing & Industrial (20)
        "caterpillar.com", "johnsoncontrols.com", "3m.com", "honeywell.com", "siemens.com", "ge.com", "emerson.com", "schneider-electric.com", "rockwellautomation.com", "abb.com", "bosch.com", "hitachihightech.com", "daikin.com", "cummins.com", "whirlpoolcorp.com", "jcb.com", "doosan.com", "yamaha-motor.com", "unitedtechnologies.com", "raytheon.com",

        // 25. Real Estate (20)
        "zillow.com", "realtor.com", "redfin.com", "trulia.com", "homes.com", "remax.com", "century21.com", "coldwellbanker.com", "kw.com", "sothebysrealty.com", "compass.com", "corcoran.com", "zillowgroup.com", "loopnet.com", "officespace.com", "costar.com", "cushmanwakefield.com", "jll.com", "savills.com", "colliers.com"
    ]);

    const personalDomains = new Set([
        "gmail.com", "googlemail.com", "outlook.com", "hotmail.com", "live.com", "msn.com",
        "hotmail.co.uk", "live.ca", "yahoo.com", "yahoo.co.uk", "yahoo.co.in", "ymail.com",
        "rocketmail.com", "icloud.com", "me.com", "mac.com", "aol.com", "verizon.net", "zoho.com",
        "mail.com", "consultant.com", "email.com", "usa.com", "post.com", "dr.com",
        "protonmail.com", "proton.me", "tutanota.com", "tutanota.de", "gmx.com", "gmx.de",
        "fastmail.com", "fastmail.fm", "messagingengine.com", "yandex.com", "yandex.ru",
        "mailfence.com", "comcast.net", "att.net", "cox.net", "bellsouth.net", "shaw.ca",
        "rogers.com", "telus.net", "btinternet.com", "orange.fr", "wanadoo.fr", "t-online.de",
        "runbox.com", "posteo.net", "neomailbox.com", "countermail.com", "startmail.com", "lavabit.com"
    ]);

    const BADGE = (txt, title) =>
        `<span class="inline-badge" title="${title}">⚠️ ${txt}</span>`;

    window._identifyEmailVersion = "v37";

    // BEGIN v38 addition: track user's domain and internal trust
    window.__userDomain = "";
    window.__internalSenderTrusted = false;
    // END v38 addition

    /* ---------- 2. OFFICE READY ---------- */
    Office.onReady(() => {
        $(document).ready(() => {
            const banner = new components.MessageBanner(document.querySelector(".MessageBanner"));
            banner.hideBanner();
            initTheme();
            wireThemeToggle();
            wireCollapsibles();
            loadProps();
            Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, loadProps);
        });
    });

    /* ---------- 3. THEME ---------- */
    function initTheme() {
        const s = localStorage.getItem(THEME_KEY) || "light";
        setTheme(s);
        $("#themeToggle").prop("checked", s === "dark");
    }
    function wireThemeToggle() {
        $("#themeToggle").on("change", function () {
            const m = this.checked ? "dark" : "light";
            setTheme(m);
            localStorage.setItem(THEME_KEY, m);
        });
    }
    function setTheme(m) {
        $("body").toggleClass("dark-mode", m === "dark");
    }

    /* ---------- 4. COLLAPSIBLES ---------- */
    function wireCollapsibles() {
        $(document).on("click", ".card.collapsible > .section-title", function () {
            $(this).closest(".card").toggleClass("collapsed");
        });
        // clicking any flag badge expands ATTACHMENTS card
        $(document).on("click", "#attachBadgeContainer .inline-badge", () => $("#attachments-card").removeClass("collapsed"));
    }

    /* ---------- 5. MAIN LOAD ---------- */
    function loadProps() {
        const it = Office.context.mailbox.item;
        if (!it) return;

        // v38 addition: track user domain globally
        window.__userDomain = baseDom(dom(Office.context.mailbox.userProfile.emailAddress));
        window.__internalSenderTrusted = false;

        // meta
        $("#dateTimeCreated").text(it.dateTimeCreated.toLocaleString());
        $("#dateTimeModified").text(it.dateTimeModified.toLocaleString());
        $("#itemClass").text(it.itemClass);
        $("#itemId").text(it.itemId);
        $("#itemType").text(it.itemType);

        // attachments
        renderAttachments(it);

        // URLs
        $("#urls").text("Scanning…");
        scanBodyUrls(it, urls => {
            $("#urls").html(urls.length ? urls.map(shortUrlSpan).join("<br/>") : "None");

            const senderBase = baseDom(dom((it.sender?.emailAddress || it.from.emailAddress || "").toLowerCase()));
            const userBase = baseDom(dom(Office.context.mailbox.userProfile.emailAddress || ""));
            const allDomains = urls.map(u => {
                try {
                    return baseDom(new URL(u).hostname.toLowerCase());
                } catch {
                    return null;
                }
            }).filter(Boolean);

            const uniqueDomains = new Set(allDomains);
            const senderCount = allDomains.filter(d => d === senderBase).length;
            const userCount = allDomains.filter(d => d === userBase).length;
            const externalCount = urls.length - senderCount;

            const $sec = $("#securityBadgeContainer").empty();

            // add badges only when count > 0
            if (externalCount) {
                $sec.prepend(BADGE(`${externalCount} external URL${externalCount !== 1 ? "s" : ""}`, `URLs not matching sender’s domain`));
            }
            if (userCount) {
                $sec.prepend(BADGE(`${userCount} match Your Domain`, `Your domain (${userBase}) appears ${userCount} time(s)`));
            }
            if (senderCount) {
                $sec.prepend(BADGE(`${senderCount} match Sender Domain`, `Sender’s domain (${senderBase}) appears ${senderCount} time(s)`));
            }
            if (urls.length) {
                // totals only if at least 1 URL
                $sec.prepend(
                    BADGE(
                        `${urls.length} URL${urls.length !== 1 ? "s" : ""} | ${uniqueDomains.size} DOMAIN${uniqueDomains.size !== 1 ? "s" : ""}`,
                        "Total URLs and unique domains"
                    )
                );
            }

            // collapse Security Flags card if empty
            if (!$sec.children().length) {
                $("#security-card").addClass("collapsed");
            } else {
                $("#security-card").removeClass("collapsed");
            }
        });

        // addresses (with truncation helper)
        $("#from").html(formatAddr(it.from));
        $("#sender").html(formatAddr(it.sender));
        $("#to").html(formatAddrs(it.to));
        $("#cc").html(formatAddrs(it.cc));
        $("#subject").text(it.subject);

        $("#conversationId").html(truncateText(it.conversationId));
        $("#internetMessageId").html(truncateText(it.internetMessageId));
        $("#normalizedSubject").text(it.normalizedSubject);

        // Original order unchanged:
        senderClassification(it);
        checkAuthHeaders(it);
        fromSenderMismatch(it);
    }

    /* ---------- 6. ATTACHMENTS ---------- */
    function renderAttachments(it) {
        let list = it.attachments || [];
        if (list.length) {
            fill(list);
            return;
        }
        if (it.getAttachmentsAsync) {
            $("#attachments").text("Loading…");
            it.getAttachmentsAsync(r => {
                list = r.status === "succeeded" ? r.value : [];
                fill(list);
            });
        } else {
            fill([]);
        }
    }
    function fill(l) {
        $("#attachments").html(l.length ? l.map(a => truncateText(a.name, true)).join("<br/>") : "None");
        const $ac = $("#attachBadgeContainer").empty();
        if (l.length) {
            $ac.append(BADGE(`${l.length} ATTACHMENT${l.length !== 1 ? "s" : ""}`, "Review attachments before opening"));
        }
    }

    /* ---------- 7. URL HELPERS ---------- */
    function scanBodyUrls(it, cb) {
        it.body.getAsync(Office.CoercionType.Text, r => {
            if (r.status !== "succeeded") {
                cb([]);
                return;
            }
            const m = r.value.match(/https?:\/\/[^\s"'<>]+/gi) || [];
            cb([...new Set(m)].slice(0, 200));
        });
    }
    function shortUrlSpan(u) {
        const s = truncateUrl(u, 30);
        return `<span class="short-url" title="${escapeHtml(u)}">${escapeHtml(s)}</span>`;
    }
    function truncateUrl(u, max) {
        try {
            const { protocol, hostname, pathname } = new URL(u);
            const shortPath = pathname.length > max ? pathname.slice(0, max) + "…" : pathname;
            return `${protocol}//${hostname}${shortPath}`;
        } catch {
            return u.length > 60 ? u.slice(0, 57) + "…" : u;
        }
    }
    function escapeHtml(s) {
        return s.replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", "\"": "&quot;", "'": "&#39;" }[c]));
    }

    /* ---------- 8. SENDER TYPE / VERIFIED ------------- */
    function senderClassification(it) {
        const email = (it.from?.emailAddress || "").toLowerCase();
        const base = baseDom(dom(email));

        // Enhanced check: entire email in 'verifiedSenders' OR domain in 'verifiedDomains'.
        // v38 addition: also consider internalSenderTrusted
        const isVerified = verifiedSenders.includes(email) || verifiedDomains.has(base) || window.__internalSenderTrusted;

        const vCls = isVerified ? "badge-verified" : "badge-unverified";
        const personal = personalDomains.has(base);
        const cCls = personal ? "badge-personal" : "badge-business";
        const cTxt = (personal ? "⚠️ " : "") + "Sender is " + (personal ? "Personal Email" : "Business Email");

        $("#classBadgeContainer").html(`<div class='badge ${cCls}'>${cTxt}</div>`);
        $("#verifiedBadgeContainer").html(
            `<div class='badge ${vCls}'>${isVerified ? "Verified Sender" : "Not Verified"}: ${email}</div>`
        );
    }

    /* ---------- 9. AUTH HEADERS ----------------------- */
    function checkAuthHeaders(it) {
        if (!it.getAllInternetHeadersAsync) return;
        it.getAllInternetHeadersAsync(r => {
            if (r.status !== "succeeded") return;
            const hdr = r.value || "";
            const lines = hdr.split(/\r?\n/);

            let spf, dkim, dmarc, envDom = null, dkimDom = null;
            lines.forEach(l => {
                const low = l.toLowerCase();
                if (low.includes("authentication-results:") || low.includes("arc-authentication-results:")) {
                    spf ??= val(low, "spf=");
                    dkim ??= val(low, "dkim=");
                    dmarc ??= val(low, "dmarc=");
                    if (low.includes("smtp.mailfrom=")) {
                        const m = low.match(/smtp\.mailfrom=([^;\s]+)/);
                        if (m) envDom = baseDom(dom(m[1]));
                    }
                }
                if (low.startsWith("return-path:")) {
                    const m = l.match(/<([^>]+)>/);
                    if (m) envDom = baseDom(dom(m[1]));
                }
                if (low.startsWith("dkim-signature:") && !dkimDom) {
                    const mm = l.match(/\bd=([^;]+)/i);
                    if (mm) dkimDom = baseDom(mm[1].trim().toLowerCase());
                }
            });

            const summary =
                `<div class='auth-summary ${(spf === "pass" && dkim === "pass" && dmarc === "pass") ? "auth-pass" : "auth-fail"}'>
                    SPF=${spf || "N/A"} | DKIM=${dkim || "N/A"} | DMARC=${dmarc || "N/A"}
                </div>`;

            $("#authContainer").html(summary);

            const fromBase = baseDom(dom(it.from.emailAddress));
            const dispBase = baseDom(dispDomFrom(it.from.displayName));
            const mis = [];
            if (envDom && envDom !== fromBase) mis.push(`Mail‑from ${envDom}`);
            if (dkimDom && dkimDom !== fromBase) mis.push(`DKIM d=${dkimDom}`);
            if (dispBase && dispBase !== fromBase) mis.push(`Display "${dispBase}"`);

            if (mis.length) {
                $("#authContainer").prepend(
                    BADGE("DOMAIN SENDER MISMATCH", `From: ${fromBase}\nMismatched E-mail Address: ${mis.join(", ")}`)
                );
            }
            if (mis.length || (spf && spf !== "pass") || (dkim && dkim !== "pass") || (dmarc && dmarc !== "pass")) {
                $("#auth-card").removeClass("collapsed");
            }

            // v38 addition: Safest internal trust logic:
            // If fromBase and envDom both match userDomain, that domain is not personal, and SPF/DKIM/DMARC are all pass,
            // then we trust this as an internal business sender.
            if (
                window.__userDomain &&
                fromBase === window.__userDomain &&
                envDom === window.__userDomain &&
                !personalDomains.has(window.__userDomain) &&
                spf === "pass" &&
                dkim === "pass" &&
                dmarc === "pass"
            ) {
                window.__internalSenderTrusted = true;
            }

            // v39 fix: after the async check finishes, re-run classification & mismatch
            // to ensure the UI is updated if we just set __internalSenderTrusted=true
            senderClassification(it);
            fromSenderMismatch(it);
        });
    }

    /* ---------- 10. FROM vs SENDER -------------------- */
    function fromSenderMismatch(it) {
        const fromBase = baseDom(dom(it.from?.emailAddress || ""));
        const senderBase = baseDom(dom(it.sender?.emailAddress || ""));
        if (!fromBase || !senderBase || fromBase === senderBase) return;
        $("#authContainer").prepend(
            BADGE("FROM ⁄ SENDER MISMATCH", `From: ${fromBase}\nSender: ${senderBase}`)
        );
        $("#auth-card").removeClass("collapsed");
    }

    /* ---------- 11. UTIL + TRUNCATE TEXT -------------- */
    function val(s, t) {
        if (!s.includes(t)) return null;
        const parts = s.split(t);
        if (parts.length < 2) return null;
        const match = parts[1].trim().match(/^(\w+)/);
        return match ? match[1] : null;
    }
    function dom(a) {
        return a?.match(/@([A-Za-z0-9.-]+\.[A-Za-z]{2,})$/)?.[1]?.toLowerCase() || null;
    }
    function baseDom(d) {
        if (!d) return "";
        // remove leading subdomains like www, m, l, etc.
        d = d.replace(/^(?:www\d*|m\d*|l\d*)\./i, "");
        const p = d.split(".");
        return p.length <= 2 ? d : p.slice(-2).join(".");
    }
    function dispDomFrom(n) {
        return n?.match(/@([A-Za-z0-9.-]+\.[A-Za-z]{2,})/)?.[1]?.toLowerCase() || null;
    }

    function truncateText(txt, isFile = false, max = 48) {
        if (!txt) return "";
        if (txt.length <= max) return escapeHtml(txt);
        const ell = escapeHtml(txt.slice(0, max - 1) + "…");
        return `<span class="truncate" title="${escapeHtml(txt)}">${ell}</span>`;
    }
    function escapeHtml(s) {
        return s.replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", "\"": "&quot;", "'": "&#39;" }[c]));
    }
    function formatAddr(a) {
        return `${a.displayName} &lt;${a.emailAddress}&gt;`;
    }
    function formatAddrs(arr) {
        return arr?.length ? arr.map(formatAddr).join("<br/>") : "None";
    }
})();
