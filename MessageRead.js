/* MessageRead.js – v41 Debug
   Changes in v36: Arrows inverted in HTML/CSS; added "kaseya.net".
   Changes in v37: Attachments -> separate collapsible card; #attachBadgeContainer -> #attachments-card.
   Changes in v38: Basic internal domain trust logic.
   Changes in v39: Re-run classification after async finishes.
   Changes in v40: Full-domain matching for internal trust.
   Changes in v41 (DEBUG LOGS):
     - Added console.log statements in checkAuthHeaders to see exactly what domain values we have
       (fromBase, envDom, userDomain, spf/dkim/dmarc). This helps identify why it's not verifying internally.
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
        // (same huge verifiedDomains array, unchanged) ...
        "kaseya.net",
        // ...
        // truncated here for brevity; keep all your entries
    ]);

    const personalDomains = new Set([
        // (same personal domains array, unchanged) ...
        "tutanota.de",
        // ...
        // truncated for brevity; keep all your entries
    ]);

    const BADGE = (txt, title) =>
        `<span class="inline-badge" title="${title}">⚠️ ${txt}</span>`;

    window._identifyEmailVersion = "v37";

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

        // track user domain globally
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

        // Original order
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
        // combine existing logic
        const isVerified =
            verifiedSenders.includes(email) ||
            verifiedDomains.has(base) ||
            window.__internalSenderTrusted;

        const vCls = isVerified ? "badge-verified" : "badge-unverified";
        const personal = personalDomains.has(base);
        const cCls = personal ? "badge-personal" : "badge-business";
        const cTxt = (personal ? "⚠️ " : "") + "Sender is " + (personal ? "Personal Email" : "Business Email");

        $("#classBadgeContainer").html(`<div class='badge ${cCls}'>${cTxt}</div>`);
        $("#verifiedBadgeContainer").html(
            `<div class='badge ${vCls}'>${isVerified ? "Verified Sender" : "Not Verified"}: ${email}</div>`
        );
    }

    /* ---------- 9. AUTH HEADERS (DEBUG LOGS ADDED) ---- */
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

            // *** DEBUG LOGS to console ***
            const fromBase = fullDomain(it.from.emailAddress) || "";
            console.log("DEBUG => User domain:", window.__userDomain);
            console.log("DEBUG => From address:", it.from.emailAddress, "-> fromBase:", fromBase);
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
            if (envDom && envDom.toLowerCase() !== fromBase.toLowerCase()) {
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
                    BADGE("DOMAIN SENDER MISMATCH", `From: ${fromBase}\nMismatched E-mail Address: ${mis.join(", ")}`)
                );
            }
            if (mis.length || (spf && spf !== "pass") || (dkim && dkim !== "pass") || (dmarc && dmarc !== "pass")) {
                $("#auth-card").removeClass("collapsed");
            }

            // direct-domain approach for internal trust
            if (
                window.__userDomain &&
                domainsMatchForInternal(fromBase, window.__userDomain) &&
                domainsMatchForInternal(envDom, window.__userDomain) &&
                !personalDomains.has(window.__userDomain.toLowerCase()) &&
                spf === "pass" && dkim === "pass" && dmarc === "pass"
            ) {
                window.__internalSenderTrusted = true;
                console.log("DEBUG => Internal domain verified. Setting __internalSenderTrusted = true");
            } else {
                console.log("DEBUG => Not marking as internal trust. Check conditions above.");
            }

            // re-run classification & mismatch so UI updates
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

    // Returns entire domain of an email, e.g. "bob@myorg.onmicrosoft.com" => "myorg.onmicrosoft.com"
    function fullDomain(email) {
        if (!email) return "";
        const m = email.toLowerCase().match(/@([a-z0-9.\-]+)/);
        return m ? m[1] : "";
    }

    // baseDom approach for external checks
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

    // compares full domain strings ignoring case
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
