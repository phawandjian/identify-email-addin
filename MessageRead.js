/* MessageRead.js – v68
   CHANGES from v67:
   1) If “Safe” or “PossiblyNotSafe” statuses, forcibly set black text using setProperty("color","#000","important").
   2) Bumped version from 67 to 68.
*/

/* 
   Additional CHANGES for v71 (verbiage fixes):
   - Renamed "✔️ SPF PASS" to "✔️ Server Check" (when SPF is pass) 
     + new tooltip: "We confirmed that the sender is a match with the identity coming from that domain."
   - Renamed "❌ DKIM N/A" to "❌ Integrity Check" (when DKIM is none or N/A)
     + new tooltip: "We couldn't detect that the authorized domain matches the one you see."
   - Renamed "❌ DMARC N/A" to "❌ Sender Match" (when DMARC is none or N/A)
     + new tooltip: "We couldn't detect that the sender is authentic and the domain matches the brand shown in ‘From:’"
   - Authentication card header is now "Anti-Spoofing Checks" in the HTML.
   - Authentication Summary lines are each on their own row, in black text (instead of red).
*/

/*
   CHANGED in v72:
   - Bumped version from 71 to 72.
   - Added modal-based help for Anti-Spoofing Checks and Security Flags.

   CHANGED in v73:
   - Bumped version from 72 to 73 below.
   - Now we attempt to open a separate Outlook dialog window (displayDialogAsync)
     to display help info outside the task pane. If that fails or isn’t supported,
     we fall back to the in-pane overlay modal.
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
        // (list of domains remains unchanged)
        "amazon.com", "ebay.com", "alibaba.com", "aliexpress.com", "jd.com", "walmart.com", "target.com", "rakuten.com", "mercadolibre.com", "flipkart.com", "overstock.com", "etsy.com", "groupon.com", "wayfair.com", "zappos.com", "shein.com", "gearbest.com", "banggood.com", "tmall.com", "shopify.com",

        "costco.com", "kohls.com", "bestbuy.com", "macys.com", "nordstrom.com", "bloomingdales.com", "dillards.com", "jcpenney.com", "sears.com", "neimanmarcus.com", "saksfifthavenue.com", "meijer.com", "biglots.com", "rossstores.com", "tjmaxx.com", "marshalls.com", "burlington.com", "dollargeneral.com", "familydollar.com", "bedbathandbeyond.com",

        "gap.com", "oldnavy.com", "bananarepublic.com", "uniqlo.com", "hm.com", "zara.com", "forever21.com", "asos.com", "revolve.com", "urbanoutfitters.com", "freepeople.com", "anthropologie.com", "abercrombie.com", "hollisterco.com", "fashionnova.com", "victoriassecret.com", "adidas.com", "nike.com", "underarmour.com", "lululemon.com",

        "microsoft.com", "apple.com", "google.com", "oracle.com", "sap.com", "salesforce.com", "adobe.com", "ibm.com", "intel.com", "dell.com", "hp.com", "lenovo.com", "asus.com", "nvidia.com", "amd.com", "autodesk.com", "zoom.us", "slack.com", "gitlab.com", "atlassian.com",
        "kaseya.net",

        "samsung.com", "lg.com", "sony.com", "panasonic.com", "philips.com", "sharpusa.com", "huawei.com", "xiaomi.com", "oneplus.com", "realme.com", "oppo.com", "vivo.com", "toshiba.com", "pioneer.com", "jvc.com", "canon.com", "nikon.com", "epson.com", "fujifilm.com", "bose.com",

        "paypal.com", "stripe.com", "squareup.com", "venmo.com", "skrill.com", "payoneer.com", "wepay.com", "adyen.com", "authorize.net", "alipay.com", "neteller.com", "googlepay.com", "amazonpay.com", "worldpay.com", "firstdata.com", "payu.com", "bill.com", "intuit.com", "xero.com", "coinbase.com",

        "chase.com", "wellsfargo.com", "bankofamerica.com", "citi.com", "usbank.com", "pnc.com", "truist.com", "capitalone.com", "americanexpress.com", "discover.com", "goldmansachs.com", "barclays.com", "hsbc.com", "lloydsbank.com", "rbs.co.uk", "santander.com", "bbva.com", "bnymellon.com", "sofi.com", "ally.com",

        "geico.com", "progressive.com", "allstate.com", "statefarm.com", "farmers.com", "usaa.com", "libertymutual.com", "nationwide.com", "travelers.com", "chubb.com", "zurichna.com", "thehartford.com", "metlife.com", "prudential.com", "aetna.com", "cigna.com", "humana.com", "aflac.com", "coloniallife.com", "globelife.com",

        "pfizer.com", "moderna.com", "johnsonandjohnson.com", "merck.com", "astrazeneca.com", "novartis.com", "roche.com", "gsk.com", "sanofi.com", "abbvie.com", "bristolmyerssquibb.com", "lilly.com", "bayer.com", "amgen.com", "teva.com", "viatris.com", "regeneron.com", "cardinalhealth.com", "mckesson.com", "abbott.com",

        "att.com", "verizon.com", "t-mobile.com", "sprint.com", "xfinity.com", "comcast.com", "charter.com", "spectrum.com", "centurylink.com", "frontier.com", "bt.com", "vodafone.com", "orange.com", "telefonica.com", "rogers.com", "bell.ca", "telus.com", "telstra.com", "mtn.com", "uscellular.com",

        "facebook.com", "instagram.com", "twitter.com", "linkedin.com", "snapchat.com", "pinterest.com", "tiktok.com", "reddit.com", "tumblr.com", "weibo.com", "wechat.com", "discord.com", "quora.com", "meetup.com", "xing.com", "vk.com", "flickr.com", "behance.net", "deviantart.com", "medium.com",

        "baidu.com", "yandex.com", "cloudflare.com", "akamai.com", "digitalocean.com", "rackspace.com", "godaddy.com", "namecheap.com", "wordpress.com", "squarespace.com", "weebly.com", "wix.com", "bigcommerce.com", "mailchimp.com", "hubspot.com", "constantcontact.com", "webex.com", "cisco.com", "github.com", "tencent.com",

        "booking.com", "expedia.com", "tripadvisor.com", "orbitz.com", "travelocity.com", "priceline.com", "kayak.com", "skyscanner.com", "trivago.com", "hotwire.com", "hopper.com", "agoda.com", "cheapoair.com", "ebookers.com", "cheapair.com", "airfarewatchdog.com", "lastminute.com", "travelzoo.com", "travelgenio.com", "momondo.com",

        "delta.com", "united.com", "southwest.com", "american.com", "aa.com", "alaskaair.com", "jetblue.com", "spirit.com", "hawaiianairlines.com", "allegiantair.com", "britishairways.com", "lufthansa.com", "airfrance.com", "klm.com", "emirates.com", "qatarairways.com", "etihad.com", "cathaypacific.com", "singaporeair.com", "aerlingus.com",

        "marriott.com", "hilton.com", "hyatt.com", "ihg.com", "choicehotels.com", "wyndhamhotels.com", "accor.com", "ritzcarlton.com", "fourseasons.com", "fairmont.com", "starwoodhotels.com", "mgmresorts.com", "wynnresorts.com", "hostels.com", "motel6.com", "bestwestern.com", "radissonhotels.com", "scandichotels.com", "oyorooms.com", "airbnb.com",

        "hertz.com", "avis.com", "budget.com", "enterprise.com", "alamo.com", "nationalcar.com", "thrifty.com", "dollar.com", "sixt.com", "uhaul.com", "pensketruckrental.com", "lyft.com", "uber.com", "grab.com", "bolt.eu", "cabify.com", "lime.me", "bird.co", "spin.app", "turo.com",

        "starbucks.com", "dunkindonuts.com", "mcdonalds.com", "burgerking.com", "wendys.com", "tacobell.com", "pizzahut.com", "dominos.com", "papajohns.com", "chipotle.com", "panerabread.com", "chick-fil-a.com", "kfc.com", "subway.com", "fiveguys.com", "sonicdrivein.com", "arbys.com", "dairyqueen.com", "littlecaesars.com", "jimmyjohns.com",

        "ups.com", "fedex.com", "dhl.com", "usps.com", "canadapost.ca", "royalmail.com", "parcelforce.com", "hermesworld.com", "dpd.com", "tnt.com", "aramex.com", "gls-group.eu", "yamato-hd.co.jp", "japanpost.jp", "laposte.fr", "upsupplychain.com", "fedexcustomcritical.com", "dhlglobalforwarding.com", "ontrac.com", "yrc.com",

        "netflix.com", "hulu.com", "disneyplus.com", "hbo.com", "showtime.com", "paramountplus.com", "peacocktv.com", "discoveryplus.com", "espn.com", "fox.com", "abc.com", "nbc.com", "cbs.com", "bbc.co.uk", "cnn.com", "bloomberg.com", "reuters.com", "theguardian.com", "nytimes.com", "wsj.com",

        "ford.com", "gm.com", "chevrolet.com", "toyota.com", "honda.com", "nissanusa.com", "hyundaiusa.com", "kia.com", "tesla.com", "bmw.com", "mercedes-benz.com", "audi.com", "volkswagen.com", "porsche.com", "volvo.com", "subaru.com", "mazdausa.com", "dodge.com", "jeep.com", "ramtrucks.com",

        "harvard.edu", "mit.edu", "stanford.edu", "berkeley.edu", "ox.ac.uk", "cam.ac.uk", "yale.edu", "princeton.edu", "columbia.edu", "ucla.edu", "nyu.edu", "upenn.edu", "caltech.edu", "cmu.edu", "gatech.edu", "uf.edu", "umich.edu", "k12.com", "coursera.org", "edx.org",

        "un.org", "who.int", "worldbank.org", "imf.org", "wto.org", "unesco.org", "unicef.org", "redcross.org", "salvationarmy.org", "unitedway.org", "habitat.org", "wwf.org", "greenpeace.org", "amnesty.org", "doctorswithoutborders.org", "care.org", "oxfam.org", "mercycorps.org", "charitywater.org", "worldvision.org",

        "usa.gov", "irs.gov", "ssa.gov", "nps.gov", "nasa.gov", "gov.uk", "canada.ca", "australia.gov.au", "india.gov.in", "gov.cn", "europa.eu", "whitehouse.gov", "senate.gov", "house.gov", "justice.gov", "ny.gov", "ca.gov", "gov.za", "scot.gov", "uscis.gov",

        "caterpillar.com", "johnsoncontrols.com", "3m.com", "honeywell.com", "siemens.com", "ge.com", "emerson.com", "schneider-electric.com", "rockwellautomation.com", "abb.com", "bosch.com", "hitachihightech.com", "daikin.com", "cummins.com", "whirlpoolcorp.com", "jcb.com", "doosan.com", "yamaha-motor.com", "unitedtechnologies.com", "raytheon.com",

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

    // CHANGED in v73: version updated here
    window._identifyEmailVersion = "v73";

    // track user's domain and internal trust
    window.__userDomain = "";
    window.__internalSenderTrusted = false;

    // NEW: We’ll track some global info for the safety banner
    window._spfResult = null;
    window._dkimResult = null;
    window._dmarcResult = null;
    window._hasDomainMismatch = false;
    window._isVerifiedSender = false;
    window._externalUrlCount = 0;
    window._attachmentCount = 0;

    /* ---------- 2. OFFICE READY ---------- */
    Office.onReady(() => {
        $(document).ready(() => {
            const banner = new components.MessageBanner(document.querySelector(".MessageBanner"));
            banner.hideBanner();

            initTheme();
            wireThemeToggle();
            wireCollapsibles();
            initCopyButtons();

            loadProps();

            // Re-load on item changed (i.e. user selects a different message)
            Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, loadProps);

            // ---------- NEW: Fix the “X” close button so it closes the entire task pane ----------
            $(document).on("click", ".MessageBanner-close", function (evt) {
                evt.preventDefault();

                // Attempt to hide or close the add-in task pane
                if (Office && Office.addin && typeof Office.addin.hide === "function") {
                    Office.addin.hide();
                } else if (Office && Office.context && Office.context.ui && typeof Office.context.ui.closeContainer === "function") {
                    Office.context.ui.closeContainer();
                } else {
                    // Fallback if neither API is available
                    window.close();
                }
            });
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

    function isDarkMode() {
        // helper to see if user has toggled to dark
        return $("body").hasClass("dark-mode");
    }

    /* ---------- 4. COLLAPSIBLES ---------- */
    function wireCollapsibles() {
        // Expand/collapse on section title
        $(document).on("click", ".card.collapsible > .section-title", function () {
            $(this).closest(".card").toggleClass("collapsed");
        });
        // Example: clicking attachments badge expands the attachments card
        $(document).on("click", "#attachBadgeContainer .inline-badge", () => $("#attachments-card").removeClass("collapsed"));
    }

    /* ---------- 5. MAIN LOAD ---------- */
    function loadProps() {
        const it = Office.context.mailbox.item;
        if (!it) return;

        window.__userDomain = fullDomain(Office.context.mailbox.userProfile.emailAddress);
        window.__internalSenderTrusted = false;

        // Reset our tracking for new item
        window._spfResult = null;
        window._dkimResult = null;
        window._dmarcResult = null;
        window._hasDomainMismatch = false;
        window._isVerifiedSender = false;
        window._externalUrlCount = 0;
        window._attachmentCount = 0;

        // Fill item props
        $("#dateTimeCreated").text(it.dateTimeCreated.toLocaleString());
        $("#dateTimeModified").text(it.dateTimeModified.toLocaleString());
        $("#itemClass").text(it.itemClass);

        // also store full text for copy
        $("#dateTimeCreated").data("fulltext", it.dateTimeCreated.toLocaleString());
        $("#dateTimeModified").data("fulltext", it.dateTimeModified.toLocaleString());
        $("#itemClass").data("fulltext", it.itemClass);

        // Use existing truncateText for itemId so it remains truncated + tooltip
        $("#itemId").html(truncateText(it.itemId, false, 48));
        $("#itemType").text(it.itemType);

        $("#itemId").data("fulltext", it.itemId);
        $("#itemType").data("fulltext", it.itemType);

        // attachments
        renderAttachments(it);

        // urls
        $("#urls").text("Scanning…");
        scanBodyUrls(it, urls => {
            window._externalUrlCount = urls.length;
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

            const internalCount = allDomains.filter(d => isTrustedInternalLink(d)).length;
            const externalCount = urls.length - (senderCount + internalCount);

            const $sec = $("#securityBadgeContainer").empty();

            if (externalCount > 0) {
                $sec.prepend(BADGE(
                    `${externalCount} external URL${externalCount !== 1 ? "s" : ""}`,
                    `URLs not matching sender's domain or your internal domain`
                ));
            }
            if (userCount) {
                $sec.prepend(BADGE(
                    `${userCount} match Your Domain`,
                    `Your domain (${userBase}) appears ${userCount} time(s)`
                ));
            }
            if (senderCount) {
                $sec.prepend(BADGE(
                    `${senderCount} match Sender Domain`,
                    `Sender’s domain (${senderBase}) appears ${senderCount} time(s)`
                ));
            }
            if (urls.length) {
                $sec.prepend(
                    BADGE(
                        `${urls.length} URL${urls.length !== 1 ? "s" : ""} | ${uniqueDomains.size} DOMAIN${uniqueDomains.size !== 1 ? "s" : ""}`,
                        "Total URLs and unique domains"
                    )
                );
            }

            if (!$sec.children().length) {
                $("#security-card").addClass("collapsed");
            } else {
                $("#security-card").removeClass("collapsed");
            }

            // If checkAuthHeaders has finished by now, update banner again:
            updateEmailSafetyBanner();
        });

        // Truncate these fields
        $("#from").html(truncateText(formatAddr(it.from), false, 50));
        $("#sender").html(truncateText(formatAddr(it.sender), false, 50));
        $("#to").html(formatAddrsTruncated(it.to, 30));
        $("#cc").html(formatAddrsTruncated(it.cc, 30));
        $("#subject").html(truncateText(it.subject, false, 60));

        $("#from").data("fulltext", formatAddr(it.from));
        $("#sender").data("fulltext", formatAddr(it.sender));
        $("#to").data("fulltext", (it.to || []).map(a => formatAddr(a)).join("; "));
        $("#cc").data("fulltext", (it.cc || []).map(a => formatAddr(a)).join("; "));
        $("#subject").data("fulltext", it.subject || "");

        $("#conversationId").html(truncateText(it.conversationId));
        $("#internetMessageId").html(truncateText(it.internetMessageId));
        $("#normalizedSubject").text(it.normalizedSubject);

        $("#conversationId").data("fulltext", it.conversationId);
        $("#internetMessageId").data("fulltext", it.internetMessageId);
        $("#normalizedSubject").data("fulltext", it.normalizedSubject);

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
        $("#attachments").html(l.length ? l.map(a => truncateText(a.name, true, 48)).join("<br/>") : "None");
        window._attachmentCount = l.length;
        const $ac = $("#attachBadgeContainer").empty();
        if (l.length) {
            $ac.append(BADGE(`${l.length} ATTACHMENT${l.length !== 1 ? "s" : ""}`, "Review attachments before opening"));
        }
    }

    /* ---------- 7. URL HELPERS ---------- */
    function scanBodyUrls(it, cb) {
        it.body.getAsync(Office.CoercionType.Text, r => {
            if (r.status !== "succeeded") {
                cb([]);
                return;
            }
            const text = r.value;
            const matches = text.match(/https?:\/\/[^\s"'<>]+/gi) || [];
            const decoded = matches.map(u => decodeUrlWrappers(u));
            cb([...new Set(decoded)].slice(0, 200));
        });
    }

    function decodeUrlWrappers(originalUrl) {
        let url = originalUrl.trim();
        while (true) {
            const newUrl = decodeOnePass(url);
            if (newUrl === url) {
                break;
            }
            url = newUrl;
        }
        return url;
    }

    function decodeOnePass(originalUrl) {
        let url = originalUrl.trim();
        try {
            const lower = url.toLowerCase();

            // 1) Microsoft Safe Links
            if (lower.includes("safelinks.protection.outlook.com/") && lower.includes("?url=")) {
                const match = url.match(/[?&]url=([^&]+)/i);
                if (match && match[1]) {
                    const decodedParam = decodeURIComponent(match[1]);
                    return decodedParam.trim() || originalUrl;
                }
            }

            // 2) Proofpoint older style
            if (lower.includes("urldefense.proofpoint.com") && lower.includes("?u=")) {
                const match = url.match(/[?&]u=([^&]+)/i);
                if (match && match[1]) {
                    let decodedParam = match[1].replace(/-/g, '%');
                    try {
                        decodedParam = decodeURIComponent(decodedParam);
                        return decodedParam.trim() || originalUrl;
                    } catch {
                        // fallback
                    }
                }
            }

            // 2b) Proofpoint v3 "v3/__https://"
            if (lower.includes("urldefense.com/v3/__https://")) {
                const match = url.match(/\/v3\/__https?:\/\/(.+)/i);
                if (match && match[1]) {
                    return "https://" + match[1];
                }
            }

            // 2c) Additional Proofpoint variants (v2, v4, etc.)
            if (/urldefense\.com\/v\d+\/__http/i.test(lower)) {
                const m = url.match(/\/v(\d+)\/__http(s?):\/\/(.+)/i);
                if (m && m[3]) {
                    let proto = "http" + (m[2] || "");
                    let remainder = m[3];
                    if (remainder.includes("-")) {
                        let replaced = remainder.replace(/-/g, '%');
                        try {
                            replaced = decodeURIComponent(replaced);
                            remainder = replaced.trim() || remainder;
                        } catch { }
                    }
                    return proto + "://" + remainder;
                }
            }

            // 2e) Partial slash fix
            if (/urldefense\.com\/v\d+\/__http(s?):\/[^\s]/i.test(lower)) {
                const m = url.match(/(\/v\d+\/__http(s?):)\/([^]+)/i);
                if (m && m[3]) {
                    let proto = "http" + (m[2] || "");
                    let remainder = m[3];
                    if (remainder.includes("-")) {
                        let replaced = remainder.replace(/-/g, '%');
                        try {
                            replaced = decodeURIComponent(replaced);
                            remainder = replaced.trim() || remainder;
                        } catch { }
                    }
                    return `${proto}://${remainder}`;
                }
            }

            // 3) Symantec / ClickTime
            if (lower.includes("clicktime.symantec.com") && lower.includes("?u=")) {
                const match = url.match(/[?&]u=([^&]+)/i);
                if (match && match[1]) {
                    const decodedParam = decodeURIComponent(match[1]);
                    return decodedParam.trim() || originalUrl;
                }
            }

            // 4) aka.ms / MS learn
            if ((lower.includes("aka.ms/") || lower.includes("learn.microsoft.com")) &&
                (lower.includes("targeturl=") || lower.includes("target="))) {
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

    function isTrustedInternalLink(domain) {
        if (!domain) return false;
        if (window.__userDomain && domainsMatchForInternal(domain, window.__userDomain)) {
            return true;
        }
        return false;
    }

    function shortUrlSpan(u) {
        const s = truncateUrl(u, 30);
        let domain = "";
        try {
            domain = baseDom(new URL(u).hostname.toLowerCase());
        } catch { }

        if (isTrustedInternalLink(domain)) {
            return `<span class="short-url" title="Trusted internal link (domain matches your org)">✔️ ${escapeHtml(s)}</span>`;
        } else {
            return `<span class="short-url" title="${escapeHtml(u)}">${escapeHtml(s)}</span>`;
        }
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

        const isVerified =
            verifiedSenders.includes(email) ||
            verifiedDomains.has(base) ||
            window.__internalSenderTrusted;

        window._isVerifiedSender = isVerified;

        const vCls = isVerified ? "badge-verified" : "badge-unverified";
        const personal = personalDomains.has(base);
        const cCls = personal ? "badge-personal" : "badge-business";
        const cTxt = (personal ? "⚠️ " : "") + "Sender is " + (personal ? "Personal Email" : "Business Email");

        $("#classBadgeContainer").html(`<div class='badge ${cCls}'>${cTxt}</div>`);

        // Truncate the displayed email for Verified Sender (limit 40)
        const truncatedEmail = truncateText(email, false, 40);
        $("#verifiedBadgeContainer").html(
            `<div class='badge ${vCls}' style="white-space: normal;">
                ${isVerified ? "Verified Sender:" : "Not Verified:"}<br/>
                ${truncatedEmail}
            </div>`
        );
    }

    /* ---------- 9. AUTH HEADERS (with updated hover text) ---------- */

    function buildSpfBadge(status) {
        const s = (status || "").toLowerCase();
        let icon, cls;
        let hoverText;

        if (!status || s === "n/a" || s === "none") {
            icon = "❌";
            cls = "badge-spf-warn";
            hoverText = "No SPF record found — can’t confirm if the sender is genuine.";
        } else if (s === "pass") {
            icon = "✔️";
            cls = "badge-spf-pass";
            hoverText = "We confirmed that the sender is a match with the identity coming from that domain.";
        } else if (s === "internal") {
            icon = "✔️";
            cls = "badge-spf-pass";
            hoverText = "Internal email (SPF checks skipped).";
        } else {
            icon = "⚠️";
            cls = "badge-spf-fail";
            hoverText = "Sender not verified — the address may be spoofed.";
        }

        let label = status ? status.toUpperCase() : "N/A";
        if (s === "pass") {
            // rename text to "Server Check" for SPF PASS
            label = "Server Check";
        }

        return `<div class="badge ${cls}" title="${hoverText}">${icon}&nbsp;${label}</div>`;
    }

    function buildDkimBadge(status) {
        const s = (status || "").toLowerCase();
        let icon, cls;
        let hoverText;

        if (!status || s === "n/a" || s === "none") {
            icon = "❌";
            cls = "badge-dkim-warn";
            hoverText = "We couldn't detect that the authorized domain matches the one you see.";
        } else if (s === "pass") {
            icon = "✔️";
            cls = "badge-dkim-pass";
            hoverText = "Signature verified — the message is intact and really came from that domain.";
        } else if (s === "internal") {
            icon = "✔️";
            cls = "badge-dkim-pass";
            hoverText = "Internal email (DKIM checks skipped).";
        } else {
            icon = "⚠️";
            cls = "badge-dkim-fail";
            hoverText = "Signature invalid — the sender can’t be verified; the email may be forged or altered.";
        }

        let label = status ? status.toUpperCase() : "N/A";
        if (!status || s === "n/a" || s === "none") {
            // rename text to "Integrity Check"
            label = "Integrity Check";
        }

        return `<div class="badge ${cls}" title="${hoverText}">${icon}&nbsp;${label}</div>`;
    }

    function buildDmarcBadge(status) {
        const s = (status || "").toLowerCase();
        let icon, cls;
        let hoverText;

        if (!status || s === "n/a" || s === "none") {
            icon = "❌";
            cls = "badge-dmarc-warn";
            hoverText = "We couldn't detect that the sender is authentic and the domain matches the brand shown in ‘From:’";
        } else if (s === "pass") {
            icon = "✔️";
            cls = "badge-dmarc-pass";
            hoverText = "Policy verified — the domain approves this email.";
        } else if (s === "internal") {
            icon = "✔️";
            cls = "badge-dmarc-pass";
            hoverText = "Internal email (DMARC checks skipped).";
        } else {
            icon = "⚠️";
            cls = "badge-dmarc-fail";
            hoverText = "Policy failed — the domain rejects or can’t validate this email; treat with caution.";
        }

        let label = status ? status.toUpperCase() : "N/A";
        if (!status || s === "n/a" || s === "none") {
            // rename text to "Sender Match"
            label = "Sender Match";
        }

        return `<div class="badge ${cls}" title="${hoverText}">${icon}&nbsp;${label}</div>`;
    }

    function checkAuthHeaders(it) {
        // If purely internal (From=Sender=User domain), skip SPF/DKIM/DMARC checks and mark them as “internal”
        const fromEmail = (it.from?.emailAddress || "").toLowerCase();
        const senderEmail = (it.sender?.emailAddress || "").toLowerCase();
        if (
            fromEmail && senderEmail &&
            fromEmail === senderEmail &&
            window.__userDomain &&
            domainsMatchForInternal(fullDomain(fromEmail), window.__userDomain) &&
            domainsMatchForInternal(fullDomain(senderEmail), window.__userDomain) &&
            !personalDomains.has(window.__userDomain.toLowerCase())
        ) {
            // purely internal
            console.log("Detected internal from domain => skipping spf/dkim/dmarc checks");
            window._spfResult = "internal";
            window._dkimResult = "internal";
            window._dmarcResult = "internal";
            window.__internalSenderTrusted = true;

            senderClassification(it);
            fromSenderMismatch(it);
            updateEmailSafetyBanner();
            return;
        }

        // Otherwise, do the normal check
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

            window._spfResult = spf?.toLowerCase() || "none";
            window._dkimResult = dkim?.toLowerCase() || "none";
            window._dmarcResult = dmarc?.toLowerCase() || "none";

            const $auth = $("#authContainer").empty();
            $auth.append(buildSpfBadge(spf));
            $auth.append(buildDkimBadge(dkim));
            $auth.append(buildDmarcBadge(dmarc));

            const spfVal = spf ? spf.toUpperCase() : "N/A";
            const dkimVal = dkim ? dkim.toUpperCase() : "N/A";
            const dmarcVal = dmarc ? dmarc.toUpperCase() : "N/A";
            const isAllPass = (spfVal === "PASS" && dkimVal === "PASS" && dmarcVal === "PASS");

            const summaryCls = isAllPass ? "auth-pass" : "auth-fail";
            const summary = `
                <div style="margin-top:8px;"></div>
                <div class="auth-summary ${summaryCls}" style="color:#000;">
                    <strong>Authentication Summary:</strong><br/>
                    SPF: ${spfVal}<br/>
                    DKIM: ${dkimVal}<br/>
                    DMARC: ${dmarcVal}
                </div>
            `;
            $auth.append(summary);

            const shortFromBase = baseDom(dom(it.from.emailAddress));
            const dispBase = baseDom(dispDomFrom(it.from.displayName));
            const fromBaseFull = shortFromBase;

            const mismatches = [];
            if (envDom && envDom.toLowerCase() !== fromBaseFull.toLowerCase()) {
                mismatches.push(`Mail‑from ${envDom}`);
            }
            if (dkimDom && dkimDom !== shortFromBase) {
                mismatches.push(`DKIM d=${dkimDom}`);
            }
            if (dispBase && dispBase !== shortFromBase) {
                mismatches.push(`Display "${dispBase}"`);
            }

            if (mismatches.length) {
                $("#securityBadgeContainer").prepend(
                    BADGE("DOMAIN SENDER MISMATCH", `From: ${fromBaseFull}\nMismatched E-mail Address: ${mismatches.join(", ")}`)
                );
                $("#security-card").removeClass("collapsed");
                window._hasDomainMismatch = true;
            }

            if (
                mismatches.length ||
                (spf && spf !== "pass") ||
                (dkim && dkim !== "pass") ||
                (dmarc && dmarc !== "pass")
            ) {
                $("#auth-card").removeClass("collapsed");
            }

            // internal trust logic
            if (
                window.__userDomain &&
                domainsMatchForInternal(fromBaseFull, window.__userDomain) &&
                domainsMatchForInternal(envDom, window.__userDomain) &&
                !personalDomains.has(window.__userDomain.toLowerCase()) &&
                spf === "pass" && dkim === "pass" && dmarc === "pass"
            ) {
                window.__internalSenderTrusted = true;
            } else {
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
                }
            }

            senderClassification(it);
            fromSenderMismatch(it);

            // Final step: now that we have SPF/DKIM/DMARC, update the safety banner
            updateEmailSafetyBanner();
        });
    }

    /* ---------- 10. FROM vs SENDER -------------------- */
    function fromSenderMismatch(it) {
        const fromBase = baseDom(dom(it.from?.emailAddress || ""));
        const senderBase = baseDom(dom(it.sender?.emailAddress || ""));
        if (!fromBase || !senderBase || fromBase === senderBase) return;
        $("#securityBadgeContainer").prepend(
            BADGE("FROM ⁄ SENDER MISMATCH", `From: ${fromBase}\nSender: ${senderBase}`)
        );
        $("#security-card").removeClass("collapsed");
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

    function truncateText(txt, isFile = false, max = 48) {
        if (!txt) return "";
        if (txt.length <= max) return escapeHtml(txt);
        const ell = escapeHtml(txt.slice(0, max - 1) + "…");
        return `<span class="truncate" title="${escapeHtml(txt)}">${ell}</span>`;
    }

    function formatAddr(a) {
        return `${a.displayName} <${a.emailAddress}>`;
    }

    function formatAddrsTruncated(arr, maxLimit) {
        if (!arr || !arr.length) return "None";
        return arr.map(a => truncateText(formatAddr(a), false, maxLimit)).join("<br/>");
    }

    function initCopyButtons() {
        $(document).on("click", ".copy-btn", function (e) {
            e.preventDefault();
            const targetId = $(this).data("copyTarget");
            const $targetEl = $("#" + targetId);

            // prefer the stored full text if present
            const dataFull = $targetEl.data("fulltext");
            const textToCopy = dataFull || $targetEl.text().trim();

            if (!textToCopy) return;

            if (navigator.clipboard && navigator.clipboard.writeText) {
                navigator.clipboard.writeText(textToCopy).then(() => {
                    console.log("Copied to clipboard:", textToCopy);
                }).catch(() => {
                    console.warn("Clipboard copy failed");
                });
            } else {
                try {
                    const temp = document.createElement("textarea");
                    temp.value = textToCopy;
                    document.body.appendChild(temp);
                    temp.select();
                    document.execCommand("copy");
                    document.body.removeChild(temp);
                    console.log("Copied to clipboard via fallback:", textToCopy);
                } catch (ex) {
                    console.warn("Clipboard copy fallback failed", ex);
                }
            }
        });
    }

    /* ---------- 12. SAFETY BANNER LOGIC ---------- */
    function updateEmailSafetyBanner() {
        const spf = window._spfResult;
        const dkim = window._dkimResult;
        const dmarc = window._dmarcResult;
        const mismatch = window._hasDomainMismatch;
        const verifiedSender = window._isVerifiedSender;
        const externalUrls = window._externalUrlCount;
        const attachCount = window._attachmentCount;

        // Decide final safety
        const status = computeEmailSafety(spf, dkim, dmarc, verifiedSender, mismatch, externalUrls, attachCount);
        const bannerEl = document.getElementById("safetyBanner");

        if (!bannerEl) return; // no banner container => do nothing

        if (status === "Safe") {
            bannerEl.style.backgroundColor = "#c8f7c5"; // a light green
            bannerEl.style.setProperty("color", "#000", "important");
            bannerEl.textContent = "✅ Safe – All trust checks passed";
        } else if (status === "PossiblyNotSafe") {
            const cautionColor = isDarkMode() ? "#5E4E1C" : "#FFF4CF";
            bannerEl.style.backgroundColor = cautionColor;
            bannerEl.style.setProperty("color", "#000", "important");
            if (window.__internalSenderTrusted) {
                bannerEl.textContent = "⚠️ Likely Safe (internal), but use caution – One or more checks failed";
            } else {
                bannerEl.textContent = "⚠️ Caution – One or more checks failed";
            }
        } else {
            bannerEl.style.backgroundColor = "#f6989d"; // a softer red
            bannerEl.textContent = "❌ Unsafe – Clear indicators of risk";
        }
    }

    function computeEmailSafety(spf, dkim, dmarc, verified, mismatch, urlCount, attachCount) {
        // If spf/dkim/dmarc are "internal", treat them as pass. Then degrade to PossiblyNotSafe if links or attachments.
        if (spf === "internal" && dkim === "internal" && dmarc === "internal") {
            if (urlCount > 0 || attachCount > 0) {
                return "PossiblyNotSafe";
            } else {
                return "Safe";
            }
        }

        // 1) If SPF/DKIM/DMARC any is "fail", or domain mismatch => "Unsafe"
        // 2) If not verified, or some are "none"/"n/a", => "PossiblyNotSafe"
        // 3) If all pass + verified + no mismatch => "Safe"

        const spfFail = (spf === "fail");
        const dkimFail = (dkim === "fail");
        const dmarcFail = (dmarc === "fail");
        if (spfFail || dkimFail || dmarcFail || mismatch) {
            return "Unsafe";
        }

        const spfPass = (spf === "pass");
        const dkimPass = (dkim === "pass");
        const dmarcPass = (dmarc === "pass");
        const allPass = spfPass && dkimPass && dmarcPass;

        if (allPass && verified) {
            return "Safe";
        }

        return "PossiblyNotSafe";
    }

    /* ---------- 13. CHANGED in v73: Attempt to open help outside the pane; fallback to modal ---------- */

    window.showHelpAuth = function () {
        // We prefer an external dialog if available, otherwise fallback to the in-pane overlay
        tryDisplayDialogAsync("auth");
    };

    window.showHelpSecurity = function () {
        tryDisplayDialogAsync("security");
    };

    function tryDisplayDialogAsync(topic) {
        if (
            Office.context &&
            Office.context.ui &&
            typeof Office.context.ui.displayDialogAsync === "function"
        ) {
            // Construct minimal HTML for outside pop-up. We'll embed the same text from our in-pane help.
            let helpContentHtml = "";
            if (topic === "auth") {
                helpContentHtml = `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8"/>
  <title>Anti-Spoofing Checks</title>
</head>
<body style="font-family: sans-serif; margin: 16px;">
  <h2>Anti-Spoofing Checks</h2>
  <p><strong>SPF (Server Policy Framework)</strong>: Verifies the sending server is allowed to send on behalf of that domain.</p>
  <p><strong>DKIM (DomainKeys Identified Mail)</strong>: Ensures the message was not altered in transit and is signed by the domain’s authorized key.</p>
  <p><strong>DMARC (Domain-based Message Authentication, Reporting &amp; Conformance)</strong>: Aligns both SPF and DKIM and declares how to handle failing emails.</p>
  <p>Not all domains implement these checks yet, but their absence can be a red flag. As more providers adopt them,
  missing or failing checks can indicate spoofing or forgery.</p>
</body>
</html>
`;
            } else {
                helpContentHtml = `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8"/>
  <title>Security Flags</title>
</head>
<body style="font-family: sans-serif; margin: 16px;">
  <h2>Security Flags</h2>
  <p>This card highlights potential risks in an email, such as suspicious links, attachments, and domain mismatches.</p>
  <ul>
    <li><strong>Links</strong>: We scan all URLs. External links (not matching your organization or the sender’s domain) are flagged.</li>
    <li><strong>Attachments</strong>: Attachments can carry malware or harmful content. Always review them carefully.</li>
    <li><strong>Internal vs External Domains</strong>: We compare domains to your own and to known trusted senders. Emails from unexpected external domains may be riskier.</li>
  </ul>
  <p>Review these flags before interacting with any links or attachments you didn’t expect.</p>
</body>
</html>
`;
            }

            // Prepare a data URL so we don't rely on an external hosting page
            const encoded = btoa(unescape(encodeURIComponent(helpContentHtml)));
            const dataUrl = "data:text/html;base64," + encoded;

            Office.context.ui.displayDialogAsync(
                dataUrl,
                { width: 50, height: 60, displayInIframe: true }, // reasonable size
                function (asyncResult) {
                    if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
                        // Fallback if dialog fails
                        showHelpModalInternal(topic);
                    }
                }
            );
        } else {
            // Fallback if the API isn't available
            showHelpModalInternal(topic);
        }
    }

    function showHelpModalInternal(topic) {
        if (topic === "auth") {
            const helpHtml = `
                <h2>Anti-Spoofing Checks</h2>
                <p><strong>SPF (Server Policy Framework)</strong>: Verifies the sending server is allowed to send on behalf of that domain.</p>
                <p><strong>DKIM (DomainKeys Identified Mail)</strong>: Ensures the message was not altered in transit and is signed by the domain’s authorized key.</p>
                <p><strong>DMARC (Domain-based Message Authentication, Reporting &amp; Conformance)</strong>: Aligns both SPF and DKIM and declares how to handle failing emails.</p>
                <p>Not all domains implement these checks yet, but their absence can be a red flag. As more providers adopt them,
                missing or failing checks can indicate spoofing or forgery.</p>
            `;
            showHelpModal(helpHtml);
        } else {
            const helpHtml = `
                <h2>Security Flags</h2>
                <p>This card highlights potential risks in an email, such as suspicious links, attachments, and domain mismatches.</p>
                <ul>
                    <li><strong>Links</strong>: We scan all URLs. External links (not matching your organization or the sender’s domain) are flagged.</li>
                    <li><strong>Attachments</strong>: Attachments can carry malware or harmful content. Always review them carefully.</li>
                    <li><strong>Internal vs External Domains</strong>: We compare domains to your own and to known trusted senders. Emails from unexpected external domains may be riskier.</li>
                </ul>
                <p>Review these flags before interacting with any links or attachments you didn’t expect.</p>
            `;
            showHelpModal(helpHtml);
        }
    }

    // the same in-pane modal functions from v72 remain
    window.showHelpModal = function (content) {
        const overlay = document.getElementById("helpModalOverlay");
        const body = document.getElementById("helpModalBody");
        if (!overlay || !body) return;
        body.innerHTML = content;
        overlay.style.display = "block";
    };

    window.closeHelpModal = function () {
        const overlay = document.getElementById("helpModalOverlay");
        if (overlay) {
            overlay.style.display = "none";
        }
    };
})();
