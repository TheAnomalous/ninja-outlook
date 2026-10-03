
    import { createOutlookSafetyController } from './outlook-safety.b19972b8.js';
    // Remove credentials left by earlier browser prototypes. No client API key is used.
    try { localStorage.removeItem('ninja_openai_key'); localStorage.removeItem('ninja_model'); } catch (e) {}
    // ----------------------------------------------------
    // State & Storage
    // ----------------------------------------------------
    const state = {
      officeReady: false,
      activeMail: null,
      draftText: '',
      isGenerating: false,
      spriteFrame: 0,
      spriteState: 'idle' // 'idle' | 'busy' | 'ready'
    };

    let pendingApproval = null;
    const safety = createOutlookSafetyController({
      office: Office,
      onInvalidated() {
        pendingApproval = null;
        document.getElementById('confirmModal').classList.remove('open');
      },
      onState(next) {
        state.activeMail = next.current;
        updateMailUI(next.current);
        renderSafety(next);
      }
    });

    // ----------------------------------------------------
    // Office.js Initialization
    // ----------------------------------------------------
    Office.onReady((info) => {
      console.log('Office.js ready, host:', info.host);
      const hostPill = document.getElementById('hostPill');
      const hostStatus = document.getElementById('hostStatus');
      const dbgHost = document.getElementById('dbgHost');
      const dbgUser = document.getElementById('dbgUser');

      if (info.host === Office.HostType.Outlook) {
        state.officeReady = true;
        safety.setReady(true);
        hostPill.classList.remove('disconnected');
        hostStatus.textContent = 'Outlook Connected';
        dbgHost.textContent = 'Outlook Classic (Edge WebView2)';

        try {
          const userProfile = Office.context.mailbox.userProfile;
          if (userProfile) dbgUser.textContent = `${userProfile.displayName || ''} <${userProfile.emailAddress || ''}>`;
        } catch(e){}

        // Probe Nested App Authentication (NAA 1.1) support in this runtime
        try {
          const isNAA = (Office.context.requirements && typeof Office.context.requirements.isSetSupported === 'function')
            ? Office.context.requirements.isSetSupported('NestedAppAuth', '1.1')
            : false;
          const dbgNAA = document.getElementById('dbgNAA');
          if (dbgNAA) {
            dbgNAA.textContent = isNAA ? 'TRUE (Supported in this host)' : 'FALSE (Not supported in this host)';
            dbgNAA.style.color = isNAA ? '#34d399' : '#f87171';
            dbgNAA.style.fontWeight = 'bold';
          }
          console.log('[Ninja] NestedAppAuth 1.1 supported:', isNAA);
        } catch(e) {
          const dbgNAA = document.getElementById('dbgNAA');
          if (dbgNAA) dbgNAA.textContent = 'Error probing: ' + e.message;
        }

        // Register ItemChanged handler so selecting a new email automatically syncs!
        if (Office.context.mailbox.addHandlerAsync) {
          Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, () => {
            loadActiveEmail();
          });
        }

        // Initial load
        loadActiveEmail();
      } else {
        safety.setReady(false);
        hostStatus.textContent = 'Browser Preview';
        dbgHost.textContent = 'Web Browser (Mock Context)';
        loadMockEmail();
      }
    });

    // ----------------------------------------------------
    // Email Context Loader
    // ----------------------------------------------------
    async function loadActiveEmail() {
      return safety.refresh();
    }

    function loadMockEmail() {
      const mock = {
        subject: "Question on FHA DTI and Student Loan Payments - Smith File",
        from: "Sarah Jenkins <sjenkins@realtypartners.com>",
        body: "Hey Adam,\n\nWorking with the Smith family on 1428 Elm St. Buyer is currently on an income-driven repayment plan showing $0/month on credit report for federal student loans. Can we use the $0 payment for qualifying on FHA, or do we have to calculate 0.5% of the total balance?\n\nThanks,\nSarah",
        mode: "read"
      };
      state.activeMail = mock;
      updateMailUI(mock);
    }

    function updateMailUI(mail) {
      const subjEl = document.getElementById('mailSubject');
      const fromEl = document.getElementById('mailFromText');
      const snippetEl = document.getElementById('mailSnippet');
      const badgeEl = document.getElementById('mailModeBadge');
      const modeLabelEl = document.getElementById('mailModeLabel');
      const dbgItem = document.getElementById('dbgItem');

      if (!mail) {
        subjEl.textContent = 'Select an email in Outlook…';
        fromEl.textContent = '—';
        snippetEl.textContent = 'Click any email in your inbox to read its context.';
        badgeEl.textContent = 'No Selection';
        dbgItem.textContent = 'None';
        return;
      }

      subjEl.textContent = mail.subject;
      fromEl.textContent = mail.from;
      snippetEl.textContent = mail.body.slice(0, 300) + (mail.body.length > 300 ? '…' : '');
      badgeEl.textContent = mail.mode === 'compose' ? 'Compose Draft' : 'Inbox Read';
      modeLabelEl.textContent = mail.mode === 'compose' ? 'OUTLOOK COMPOSE WINDOW' : 'SELECTED INBOX MESSAGE';
      dbgItem.textContent = mail.subject.slice(0, 40);

      // Auto-detect program from subject / body
      const text = (mail.subject + ' ' + mail.body).toLowerCase();
      const selProgram = document.getElementById('selProgram');
      if (text.includes('fha')) selProgram.value = 'FHA';
      else if (text.includes('va ') || text.includes('veteran')) selProgram.value = 'VA';
      else if (text.includes('usda') || text.includes('rural')) selProgram.value = 'USDA';
      else if (text.includes('mshda') || text.includes('dpa') || text.includes('down payment assistance')) selProgram.value = 'MSHDA';
      else if (text.includes('freddie') || text.includes('lpa')) selProgram.value = 'Freddie Mac';
      else if (text.includes('fannie') || text.includes('du ') || text.includes('conventional')) selProgram.value = 'Fannie Mae';
    }

    // ----------------------------------------------------
    // Animated Ninja Sprite
    // ----------------------------------------------------
    setInterval(() => {
      const sprite = document.getElementById('ninjaSprite');
      if (!sprite) return;
      
      let row = 0;
      let maxFrames = 6;
      if (state.spriteState === 'busy') { row = 7; maxFrames = 8; }
      else if (state.spriteState === 'ready') { row = 8; maxFrames = 6; }

      state.spriteFrame = (state.spriteFrame + 1) % maxFrames;
      const xOffset = -(state.spriteFrame * 48);
      const yOffset = -(row * 52);
      sprite.style.backgroundPosition = `${xOffset}px ${yOffset}px`;
    }, 200);

    // ----------------------------------------------------
    // Knowledge Base & Prompt Engine
    // ----------------------------------------------------
    const GUIDELINE_RULES = {
      FHA: {
        agency: "FHA Handbook 4000.1 (Section II.A.4.b.iv)",
        summary: "Student loans with $0 IDR payments require documentation of the IDR agreement showing $0 or calculation of 0.5% of outstanding balance.",
        details: "Under FHA Single Family Handbook 4000.1: If the credit report does not show a monthly payment, or shows $0 (such as an Income-Driven Repayment plan), the Mortgagee must obtain written documentation from the servicer showing the payment is $0 under an IDR plan, OR use 0.5% of the total loan balance as the monthly qualifying payment."
      },
      "Fannie Mae": {
        agency: "Fannie Mae Selling Guide B3-6-05",
        summary: "If a student loan is in an income-driven repayment plan and the payment on credit report is $0, $0 may be used if documented with servicer letter.",
        details: "Fannie Mae permits $0 monthly payment for qualifying if documented by student loan servicer paperwork confirming $0 monthly payment under an Income-Driven Repayment (IDR) plan."
      },
      "Freddie Mac": {
        agency: "Freddie Mac Single-Family Seller/Servicer Guide 5401.2",
        summary: "If the payment is $0 on credit report, verify the student loan terms. Documented IDR $0 payments qualify with 0 payment.",
        details: "Freddie Mac allows $0 monthly payment when documented through the borrower’s IDR plan agreement, or 0.5% of the loan balance if no payment is documented."
      },
      VA: {
        agency: "VA Lenders Handbook M26-7 Chapter 4",
        summary: "Student loans deferred 12+ months post-closing may be excluded. If not deferred, use 5% of balance divided by 12.",
        details: "VA allows excluding student loan payments if deferred at least 12 months beyond the loan closing date. Otherwise, calculate payment at 5% of the balance divided by 12, or the documented IDR payment."
      },
      MSHDA: {
        agency: "MSHDA MI Home Loan Program Guidelines",
        summary: "MSHDA Down Payment Assistance ($7,500 / $10,000 in targeted areas). 1% borrower minimum contribution required.",
        details: "Borrower must contribute a minimum of 1% of the purchase price from their own funds. DPA is a zero-percent non-amortizing second mortgage due upon sale, transfer, or refinance."
      }
    };

    // ----------------------------------------------------
    // Draft Generation Logic
    // ----------------------------------------------------
    document.getElementById('btnGenerate').addEventListener('click', async () => {
      try {
        const context = safety.captureDraftContext();
        const program = document.getElementById('selProgram').value;
        const audience = document.getElementById('selAudience').value;
        const style = document.getElementById('selStyle').value;
        const instructions = document.getElementById('txtInstructions').value.trim();
        setBusy(true, "Preparing a local template…");
        const draft = generateSmartMortgageReply(context, program, audience, style, instructions);
        document.getElementById('txtDraft').value = draft;
        state.draftText = draft;
        safety.bindDraft(draft, context);
        renderCitations([GUIDELINE_RULES[program] || GUIDELINE_RULES['FHA']]);
        setReady("Local template ready. Independently verify guidance before use.");
        showToast("Template prepared for this Outlook item.", "success");
      } catch (err) {
        setReady("Draft preparation blocked.");
        showToast(err.message, "error");
      }
    });

    function generateSmartMortgageReply(email, program, audience, style, instructions) {
      const senderName = email.from.split('<')[0].trim().split(' ')[0] || 'there';
      let body = `Hi ${senderName},\n\nThanks for sending this scenario. I'll review the applicable ${program} requirements and supporting documentation before confirming the guidance.\n\n`;
      if (instructions) body += `Additional note: ${instructions}\n\n`;
      return body + `Best regards,\nAdam Fuller\nThe Fuller Team | Mortgage One\nafuller@mortgageone.com`;
    }

    function renderCitations(citations) {
      const section = document.getElementById('citationSection');
      const list = document.getElementById('citationList');
      list.innerHTML = '';

      if (!citations || !citations.length) {
        section.style.display = 'none';
        return;
      }

      citations.forEach(c => {
        const item = document.createElement('div');
        item.className = 'citation-item';
        item.innerHTML = `
          <div class="citation-header">
            <span class="citation-agency">${c.agency}</span>
            <span>Bundled reference; verify current policy</span>
          </div>
          <div class="citation-rule">${c.details || c.summary}</div>
        `;
        list.appendChild(item);
      });

      section.style.display = 'block';
    }

    function setBusy(busy, desc) {
      state.isGenerating = busy;
      state.spriteState = busy ? 'busy' : 'idle';
      document.getElementById('btnGenerate').disabled = busy;
      document.getElementById('ninjaStatusTitle').textContent = busy ? 'Preparing a local template…' : 'Ready to draft';
      document.getElementById('ninjaStatusDesc').textContent = desc || '';
    }

    function setReady(desc) {
      state.isGenerating = false;
      state.spriteState = 'ready';
      document.getElementById('btnGenerate').disabled = false;
      document.getElementById('ninjaStatusTitle').textContent = 'Draft ready for review';
      document.getElementById('ninjaStatusDesc').textContent = desc;
    }

    // ----------------------------------------------------
    // Insert only the immutable, one-use approval verified by the production controller.
    // ----------------------------------------------------
    const btnInsert = document.getElementById('btnInsert');
    const confirmModal = document.getElementById('confirmModal');
    const modalPreview = document.getElementById('modalPreview');
    const modalDesc = document.getElementById('modalDescription');

    function escapeHtml(str) {
      return String(str || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
    }
    function renderSafety(next) {
      const banner = document.getElementById('targetBindingBanner');
      const insert = document.getElementById('btnInsert');
      const confirm = document.getElementById('btnModalConfirm');
      const text = document.getElementById('txtDraft').value.trim();
      insert.disabled = !next.canApprove || !text;
      confirm.disabled = next.inFlight || !pendingApproval;
      if (!next.draft) { banner.style.display = 'none'; return; }
      banner.style.display = 'flex';
      banner.className = 'target-binding-banner ' + (next.canApprove ? 'bound' : 'mismatch');
      const label = next.canApprove ? 'Verified current item: ' : next.draft.used ? 'Draft already used. Prepare a new draft: ' : 'Blocked until the original item is verified: ';
      banner.innerHTML = '<span>' + label + '<strong>' + escapeHtml(next.draft.context.subject || '(No Subject)') + '</strong></span>';
      insert.title = next.canApprove ? 'Review an immutable approval for this item' : next.error || 'Insertion blocked';
    }
    function updateTargetBindingUI() { renderSafety(safety.getState()); }
    document.getElementById('txtDraft').addEventListener('input', updateTargetBindingUI);

    btnInsert.addEventListener('click', () => {
      try {
        pendingApproval = safety.beginApproval({
          text: document.getElementById('txtDraft').value,
          replyAll: document.getElementById('chkReplyAll').checked
        });
        modalPreview.textContent = pendingApproval.text;
        document.getElementById('modalTargetSubject').textContent = pendingApproval.context.subject || '(No Subject)';
        const c = pendingApproval.context;
        const original = [c.from, ...c.to, ...c.cc].filter(Boolean).join(', ');
        document.getElementById('modalTargetRecipient').textContent = c.mode === 'compose'
          ? 'To: ' + (c.to.join(', ') || '(none)') + '; Cc: ' + (c.cc.join(', ') || '(none)') + '; Bcc: ' + (c.bcc.join(', ') || '(none)')
          : 'Original sender / recipients: ' + (original || '(unavailable)');
        modalDesc.textContent = c.mode === 'compose'
          ? 'The displayed text will be inserted into this compose item after a fresh content and recipient check.'
          : 'Outlook will open an editable ' + (pendingApproval.replyAll ? 'reply-all' : 'reply') + ' form. Verify the recipients in Outlook before sending.';
        confirmModal.classList.add('open');
        renderSafety(safety.getState());
      } catch (e) { pendingApproval = null; showToast(e.message, 'error'); }
    });
    document.getElementById('btnModalCancel').addEventListener('click', () => {
      pendingApproval = null;
      safety.cancelApproval();
      confirmModal.classList.remove('open');
    });
    document.getElementById('btnModalConfirm').addEventListener('click', async () => {
      const approved = pendingApproval;
      pendingApproval = null;
      confirmModal.classList.remove('open');
      if (!approved) return;
      try {
        const result = await safety.confirm(approved.id);
        showToast(result.mode === 'compose' ? 'Draft inserted into the verified Outlook item.' : 'Editable reply opened. Review recipients before sending.', 'success');
      } catch (e) { showToast('Insertion blocked: ' + e.message, 'error'); }
      renderSafety(safety.getState());
    });

    // ----------------------------------------------------
    // Tab Navigation
    // ----------------------------------------------------
    const tabButtons = document.querySelectorAll('.tab-btn');
    const views = {
      desk: document.getElementById('viewDesk'),
      guidance: document.getElementById('viewGuidance'),
      queue: document.getElementById('viewQueue'),
      settings: document.getElementById('viewSettings')
    };

    tabButtons.forEach(btn => {
      btn.addEventListener('click', () => {
        const target = btn.dataset.tab;
        tabButtons.forEach(b => b.classList.remove('active'));
        btn.classList.add('active');

        Object.keys(views).forEach(k => {
          views[k].style.display = (k === target) ? 'flex' : 'none';
        });
      });
    });

    // ----------------------------------------------------
    // Guidance Search Tab
    // ----------------------------------------------------
    document.getElementById('btnSearchGuidance').addEventListener('click', () => {
      const qProg = document.getElementById('selGuidanceProgram').value;
      const qText = document.getElementById('txtGuidanceQuestion').value.trim();
      const resBox = document.getElementById('guidanceResults');
      const resContent = document.getElementById('guidanceContent');

      let matchKey = 'FHA';
      if (qProg === 'fannie') matchKey = 'Fannie Mae';
      else if (qProg === 'freddie') matchKey = 'Freddie Mac';
      else if (qProg === 'va') matchKey = 'VA';
      else if (qProg === 'mshda') matchKey = 'MSHDA';

      const rule = GUIDELINE_RULES[matchKey];
      resContent.innerHTML = `
        <div style="font-weight:700; color:#fff; font-size:12px; margin-bottom:4px;">${rule.agency}</div>
        <div style="color:var(--text-main); margin-bottom:6px;">${rule.details}</div>
        <div style="font-size:10px; color:var(--text-subtle);">Bundled template reference. No live research was performed.</div>
      `;
      resBox.style.display = 'block';
    });

    // ----------------------------------------------------
    // Client credentials are never collected. Authentication belongs to the private backend.
    // ----------------------------------------------------

    // ----------------------------------------------------
    // Utilities & Tools
    // ----------------------------------------------------
    document.getElementById('btnRefreshMail').addEventListener('click', () => {
      loadActiveEmail();
      showToast("Email synchronized.", "success");
    });

    document.getElementById('btnCopyDraft').addEventListener('click', () => {
      const draft = document.getElementById('txtDraft').value;
      if (!draft) return;
      navigator.clipboard.writeText(draft);
      showToast("Draft copied to clipboard!", "success");
    });

    document.getElementById('btnClearDraft').addEventListener('click', () => {
      document.getElementById('txtDraft').value = '';
      state.draftText = '';
      pendingApproval = null;
      safety.clearDraft();
      document.getElementById('btnInsert').disabled = true;
      document.getElementById('citationSection').style.display = 'none';
      updateTargetBindingUI();
      setBusy(false, "Draft cleared.");
    });

    function showToast(msg, type = 'info') {
      const toast = document.getElementById('toast');
      const toastMsg = document.getElementById('toastMessage');
      toast.className = `toast show ${type}`;
      toastMsg.textContent = msg;
      setTimeout(() => {
        toast.className = 'toast';
      }, 3500);
    }
  