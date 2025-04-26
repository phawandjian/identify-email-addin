/* MessageRead.js – v48
   Builds on v46. Adds:
   1) A "wrapped" address style for Verified Sender (no truncation).
   2) A "truncate + custom tooltip" style for Detailed Message Props.
   3) A copy-to-clipboard icon on each address.
   No existing lines removed; all v46 functionality retained.
*/

(function () {
    "use strict";

    /* ---------- 1. CONSTANTS ---------- */
    const THEME_KEY = "bkEmailAddinTheme";

    // Pre-approved entire email addresses
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

        // Inserted here: "kaseya.net"
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

    window._identifyEmailVersion = "v51";

    // track user's domain and internal trust
    window.__userDomain = "";
    window.__internalSenderTrusted = false;

    /* ---------- 2. OFFICE READY ---------- */
    Office.onReady(() => {
        $(document).ready(() => {
            const banner = new components.MessageBanner(document.querySelector(".MessageBanner"));
            banner.hideBanner();

            initTheme();
            wireThemeToggle();
            wireCollapsibles();
            wireCopyIcon();            // NEW: sets up copy-to-clipboard
            wireCustomTooltips();      // NEW: sets up custom tooltip logic

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

    /* ---------- NEW: COPY ICON ---------- */
    function wireCopyIcon() {
        // On click of .copy-addr, copy the "data-full" text to clipboard
        $(document).on("click", ".copy-addr", function (e) {
            e.stopPropagation();
            const textToCopy = $(this).attr("data-full") || "";
            if (!textToCopy) return;

            // Attempt to copy
            navigator.clipboard.writeText(textToCopy).then(() => {
                // Optionally show a quick alert or console log
                console.log("Copied:", textToCopy);
            }).catch(err => {
                console.warn("Copy failed:", err);
            });
        });
    }

    /* ---------- NEW: CUSTOM TOOLTIP ---------- */
    function wireCustomTooltips() {
        // We'll create a simple hover-based tooltip for .has-tooltip
        const $tooltip = $('<div id="customTooltip" style="position:absolute; z-index:9999; background:#333; color:#fff; padding:4px 8px; border-radius:4px; font-size:12px; max-width:300px; display:none; white-space:normal;"></div>');
        $("body").append($tooltip);

        let tooltipTimer = null;
        $(document)
            .on("mouseenter", ".has-tooltip", function (evt) {
                const tipText = $(this).attr("data-tooltip");
                if (!tipText) return;

                $tooltip.text(tipText).fadeIn(150);

                // Reposition near mouse
                const x = evt.pageX + 8;
                const y = evt.pageY + 8;
                $tooltip.css({ top: y, left: x });
            })
            .on("mousemove", ".has-tooltip", function (evt) {
                // move with mouse
                const x = evt.pageX + 8;
                const y = evt.pageY + 8;
                $tooltip.css({ top: y, left: x });
            })
            .on("mouseleave", ".has-tooltip", function () {
                $tooltip.hide();
            });
    }

    /* ---------- 5. MAIN LOAD ---------- */
    function loadProps() {
        const it = Office.context.mailbox.item;
        if (!it) return;

        window.__userDomain = fullDomain(Office.context.mailbox.userProfile.emailAddress);
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
                $sec.prepend(
                    BADGE(
                        `${urls.length} URL${urls.length !== 1 ? "s" : ""} | ${uniqueDomains.size} DOMAIN${uniqueDomains.size !== 1 ? "s" : ""}`,
                        "Total URLs and unique domains"
                    )
                );
            }

            if (!$sec.children().length) {
                $("#security-card").addClass("collapsed");
            } else {
                $("#security-card").removeClass("collapsed");
            }
        });

        // addresses
        // For Verified Sender (the top card), we want "wrapped" approach:
        $("#from").html(formatAddrWrapped(it.from));
        $("#sender").html(formatAddrWrapped(it.sender));

        // For "to", "cc", etc. in Detailed Props, we want "truncate + tooltip" approach:
        $("#to").html(formatAddrsTruncated(it.to));
        $("#cc").html(formatAddrsTruncated(it.cc));

        $("#subject").text(it.subject);
        $("#conversationId").html(truncateText(it.conversationId));
        $("#internetMessageId").html(truncateText(it.internetMessageId));
        $("#normalizedSubject").text(it.normalizedSubject);

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
            const decoded = m.map(u => decodeUrlWrappers(u));
            cb([...new Set(decoded)].slice(0, 200));
        });
    }

    function decodeUrlWrappers(originalUrl) {
        let url = originalUrl.trim();
        try {
            const lower = url.toLowerCase();

            // MS Safe Links
            if (lower.includes("safelinks.protection.outlook.com/") && lower.includes("?url=")) {
                const match = url.match(/[?&]url=([^&]+)/i);
                if (match && match[1]) {
                    const decodedParam = decodeURIComponent(match[1]);
                    return decodedParam.trim() || originalUrl;
                }
            }
            // Proofpoint older
            if (lower.includes("urldefense.proofpoint.com") && lower.includes("?u=")) {
                const match = url.match(/[?&]u=([^&]+)/i);
                if (match && match[1]) {
                    let decodedParam = match[1].replace(/-/g, '%');
                    try {
                        decodedParam = decodeURIComponent(decodedParam);
                        return decodedParam.trim() || originalUrl;
                    } catch { }
                }
            }
            // Proofpoint v3
            if (lower.includes("urldefense.com/v3/__https://")) {
                const match = url.match(/\/v3\/__https?:\/\/(.+)/i);
                if (match && match[1]) {
                    return "https://" + match[1];
                }
            }
            // Symantec/ClickTime
            if (lower.includes("clicktime.symantec.com") && lower.includes("?u=")) {
                const match = url.match(/[?&]u=([^&]+)/i);
                if (match && match[1]) {
                    const decodedParam = decodeURIComponent(match[1]);
                    return decodedParam.trim() || originalUrl;
                }
            }
            // aka.ms / learn
            if ((lower.includes("aka.ms/") || lower.includes("learn.microsoft.com")) && (lower.includes("targeturl=") || lower.includes("target="))) {
                const match = url.match(/[?&](?:targeturl|target)=([^&]+)/i);
                if (match && match[1]) {
                    const decodedParam = decodeURIComponent(match[1]);
                    return decodedParam.trim() || originalUrl;
                }
            }
            return originalUrl;
        } catch {
            return originalUrl;
        }
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

    /* ---------- 8. SENDER TYPE / VERIFIED ------------- */
    function senderClassification(it) {
        const email = (it.from?.emailAddress || "").toLowerCase();
        const base = baseDom(dom(email));
        const isVerified =
            verifiedSenders.includes(email) ||
            verifiedDomains.has(base) ||
            window.__internalSenderTrusted;

        console.log("DEBUG => senderClassification: email=", email,
            "base=", base,
            "verifiedDomainsHasBase=", verifiedDomains.has(base),
            "personalDomainsHasBase=", personalDomains.has(base),
            "internalSenderTrusted=", window.__internalSenderTrusted,
            "=> final isVerified=", isVerified
        );

        const vCls = isVerified ? "badge-verified" : "badge-unverified";
        const personal = personalDomains.has(base);
        const cCls = personal ? "badge-personal" : "badge-business";
        const cTxt = (personal ? "⚠️ " : "") + "Sender is " + (personal ? "Personal Email" : "Business Email");

        $("#classBadgeContainer").html(`<div class='badge ${cCls}'>${cTxt}</div>`);
        $("#verifiedBadgeContainer").html(
            `<div class='badge ${vCls}'>${isVerified ? "Verified Sender" : "Not Verified"}: ${email}</div>`
        );
    }

    /* ---------- 9. AUTH HEADERS ---------- */
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
                        if (m) envDom = fullDomain(m[1]);
                    }
                }
                if (low.startsWith("return-path:")) {
                    const m = l.match(/<([^>]+)>/);
                    if (m) envDom = fullDomain(m[1]);
                }
                if (low.startsWith("dkim-signature:") && !dkimDom) {
                    const mm = l.match(/\bd=([^;]+)/i);
                    if (mm) dkimDom = baseDom(mm[1].trim().toLowerCase());
                }
            });

            const fromBaseFull = fullDomain(it.from.emailAddress) || "";
            console.log("DEBUG => User domain:", window.__userDomain);
            console.log("DEBUG => From address:", it.from.emailAddress, "-> fromBase:", fromBaseFull);
            console.log("DEBUG => Envelope domain (envDom):", envDom);
            console.log("DEBUG => SPF:", spf, "DKIM:", dkim, "DMARC:", dmarc);

            const summary =
                `<div class='auth-summary ${(spf === "pass" && dkim === "pass" && dmarc === "pass") ? "auth-pass" : "auth-fail"}'>
                    SPF=${spf || "N/A"} | DKIM=${dkim || "N/A"} | DMARC=${dmarc || "N/A"}
                </div>`;
            $("#authContainer").html(summary);

            const dispBase = baseDom(dispDomFrom(it.from.displayName));
            const shortFromBase = baseDom(dom(it.from.emailAddress));
            const mis = [];
            if (envDom && envDom.toLowerCase() !== fromBaseFull.toLowerCase()) {
                mis.push(`Mail‑from ${envDom}`);
            }
            if (dkimDom && dkimDom !== shortFromBase) {
                mis.push(`DKIM d=${dkimDom}`);
            }
            if (dispBase && dispBase !== shortFromBase) {
                mis.push(`Display "${dispBase}"`);
            }

            if (mis.length) {
                $("#authContainer").prepend(
                    BADGE("DOMAIN SENDER MISMATCH", `From: ${fromBaseFull}\nMismatched E-mail Address: ${mis.join(", ")}`)
                );
            }
            if (mis.length || (spf && spf !== "pass") || (dkim && dkim !== "pass") || (dmarc && dmarc !== "pass")) {
                $("#auth-card").removeClass("collapsed");
            }

            // direct-domain approach for internal trust
            if (
                window.__userDomain &&
                domainsMatchForInternal(fromBaseFull, window.__userDomain) &&
                domainsMatchForInternal(envDom, window.__userDomain) &&
                !personalDomains.has(window.__userDomain.toLowerCase()) &&
                spf === "pass" && dkim === "pass" && dmarc === "pass"
            ) {
                window.__internalSenderTrusted = true;
                console.log("DEBUG => Internal domain verified. Setting __internalSenderTrusted = true");
            } else {
                console.log("DEBUG => Not marking as internal trust. Check conditions above.");

                const noAuthData =
                    (!spf || spf === "none" || spf === "null") &&
                    (!dkim || dkim === "none") &&
                    (!dmarc || dmarc === "null");
                if (
                    window.__userDomain &&
                    domainsMatchForInternal(fromBaseFull, window.__userDomain) &&
                    (!envDom || domainsMatchForInternal(envDom, window.__userDomain)) &&
                    !personalDomains.has(window.__userDomain.toLowerCase()) &&
                    noAuthData
                ) {
                    window.__internalSenderTrusted = true;
                    console.log("DEBUG => Fallback: purely internal message w/o external checks. Marking as trusted.");
                } else {
                    console.log("DEBUG => No fallback conditions met. Still not internal trust.");
                }
            }

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

    function fullDomain(email) {
        if (!email) return "";
        const m = email.toLowerCase().match(/@([a-z0-9.\-]+)/);
        return m ? m[1] : "";
    }

    function dom(a) {
        return a?.match(/@([A-Za-z0-9.-]+\.[A-Za-z]{2,})$/)?.[1]?.toLowerCase() || null;
    }
    function baseDom(d) {
        if (!d) return "";
        d = d.replace(/^(?:www\d*|m\d*|l\d*)\./i, "");
        const p = d.split(".");
        return p.length <= 2 ? d : p.slice(-2).join(".");
    }
    function dispDomFrom(n) {
        return n?.match(/@([A-Za-z0-9.-]+\.[A-Za-z]{2,})/)?.[1]?.toLowerCase() || null;
    }

    function domainsMatchForInternal(d1, d2) {
        if (!d1 || !d2) return false;
        return d1.trim().toLowerCase() === d2.trim().toLowerCase();
    }

    // Reuse the existing "truncateText" for files, etc.
    function truncateText(txt, isFile = false, max = 48) {
        if (!txt) return "";
        if (txt.length <= max) return escapeHtml(txt);
        const ell = escapeHtml(txt.slice(0, max - 1) + "…");
        return `<span class="truncate" title="${escapeHtml(txt)}">${ell}</span>`;
    }

    function escapeHtml(s) {
        return s.replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", "\"": "&quot;", "'": "&#39;" }[c]));
    }

    /* ---------- NEW: WRAPPED vs. TRUNCATED ADDRESSES ---------- */

    // "formatAddrWrapped": uses normal wrapping, no ellipsis, with copy icon
    function formatAddrWrapped(a) {
        if (!a) return "";
        const full = `${a.displayName} <${a.emailAddress}>`;
        // use normal wrapping
        return `<span style="white-space:normal; display:inline-block;">${escapeHtml(full)}</span>
                <span class="copy-addr" data-full="${escapeHtml(full)}" style="cursor:pointer; margin-left:6px;">📋</span>`;
    }
    function formatAddrsWrapped(arr) {
        if (!arr || !arr.length) return "None";
        return arr.map(a => formatAddrWrapped(a)).join("<br/>");
    }

    // "formatAddrTruncated": uses custom tooltip + ellipsis
    function formatAddrTruncated(a) {
        if (!a) return "";
        const full = `${a.displayName} <${a.emailAddress}>`;
        // We'll add .has-tooltip with data-tooltip for a custom tooltip.
        // Also add .truncate for ellipsis, using a fixed width if you prefer.
        return `<span class="truncate has-tooltip" style="max-width:200px; display:inline-block; overflow:hidden; text-overflow:ellipsis; white-space:nowrap;"
                     data-tooltip="${escapeHtml(full)}">
                    ${escapeHtml(full)}
                </span>
                <span class="copy-addr" data-full="${escapeHtml(full)}" style="cursor:pointer; margin-left:6px;">📋</span>`;
    }
    function formatAddrsTruncated(arr) {
        if (!arr || !arr.length) return "None";
        return arr.map(a => formatAddrTruncated(a)).join("<br/>");
    }
})();
