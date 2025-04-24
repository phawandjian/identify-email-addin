/* MessageRead.js – v37
   Changes:
     • Reordered cards in HTML (Verified > Type > Security > Attachments > Links > Auth > Detailed Props > Item Props)
     • Separated “attachments” and “links” into their own collapsible cards
     • Moved link-based badges out of “security-card” into “links-card”
     • No red border in .inline-badge
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
        // omitted for brevity, same as before:
        // ...
        "microsoft.com",
        "kaseya.net", // Already added in prior version
        // ...
    ]);

    const personalDomains = new Set([
        "gmail.com", "googlemail.com", "outlook.com", "hotmail.com", "live.com", "msn.com",
        // ...
        "yahoo.com", "yandex.com", "icloud.com", "me.com",
        // ...
        "lavabit.com"
    ]);

    const BADGE = (txt, title) =>
        `<span class="inline-badge" title="${title}">⚠️ ${txt}</span>`;

    window._identifyEmailVersion = "v37";

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
        // For badges inside the new “links” or “attachments” cards, auto-expand that card if user clicks a badge, if desired.
        // Example: $(document).on("click", "#linksBadgeContainer .inline-badge", () => $("#links-card").removeClass("collapsed"));
        // (Currently omitted, add if you want that behavior.)
    }

    /* ---------- 5. MAIN LOAD ---------- */
    function loadProps() {
        const it = Office.context.mailbox.item;
        if (!it) return;

        // meta
        $("#dateTimeCreated").text(it.dateTimeCreated.toLocaleString());
        $("#dateTimeModified").text(it.dateTimeModified.toLocaleString());
        $("#itemClass").text(it.itemClass);
        $("#itemId").text(it.itemId);
        $("#itemType").text(it.itemType);

        // attachments
        renderAttachments(it);

        // URLs
        $("#links").text("Scanning…");
        scanBodyUrls(it, urls => {
            $("#links").html(urls.length ? urls.map(shortUrlSpan).join("<br/>") : "None");

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

            // place link-based badges in #linksBadgeContainer
            const $lnk = $("#linksBadgeContainer").empty();

            // add badges only when count > 0
            if (externalCount) {
                $lnk.prepend(
                    BADGE(`${externalCount} external URL${externalCount !== 1 ? "s" : ""}`, "URLs not matching sender’s domain")
                );
            }
            if (userCount) {
                $lnk.prepend(
                    BADGE(`${userCount} match Your Domain`, `Your domain (${userBase}) appears ${userCount} time(s)`)
                );
            }
            if (senderCount) {
                $lnk.prepend(
                    BADGE(`${senderCount} match Sender Domain`, `Sender’s domain (${senderBase}) appears ${senderCount} time(s)`)
                );
            }
            if (urls.length) {
                // totals only if at least 1 URL
                $lnk.prepend(
                    BADGE(
                        `${urls.length} URL${urls.length !== 1 ? "s" : ""} | ${uniqueDomains.size} DOMAIN${uniqueDomains.size !== 1 ? "s" : ""}`,
                        "Total URLs and unique domains"
                    )
                );
            }

            // collapse the “links-card” if no link badges
            if (!$lnk.children().length) {
                $("#links-card").addClass("collapsed");
            } else {
                $("#links-card").removeClass("collapsed");
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
        const $ac = $("#attachmentBadgeContainer").empty();
        if (l.length) {
            $ac.append(
                BADGE(`${l.length} ATTACHMENT${l.length !== 1 ? "s" : ""}`, "Review attachments before opening")
            );
            $("#attachments-card").removeClass("collapsed");
        } else {
            $("#attachments-card").addClass("collapsed");
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
            // limit to first 200 just for performance
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
        const isVerified = verifiedSenders.includes(email) || verifiedDomains.has(base);

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
                    const m2 = l.match(/<([^>]+)>/);
                    if (m2) envDom = baseDom(dom(m2[1]));
                }
                if (low.startsWith("dkim-signature:") && !dkimDom) {
                    const mm = l.match(/\bd=([^;]+)/i);
                    if (mm) dkimDom = baseDom(mm[1].trim().toLowerCase());
                }
            });

            const summary = `
                <div class='auth-summary ${spf === "pass" && dkim === "pass" && dmarc === "pass" ? "auth-pass" : "auth-fail"}'>
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
    function formatAddr(a) {
        return `${a.displayName} &lt;${a.emailAddress}&gt;`;
    }
    function formatAddrs(arr) {
        return arr?.length ? arr.map(formatAddr).join("<br/>") : "None";
    }
    function escapeHtml(s) {
        return s.replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
    }
})();
