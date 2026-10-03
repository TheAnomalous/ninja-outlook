/* Shared production boundary for approved, current-item-only Outlook writes.
 * Office does not expose an atomic compare-and-write API. We read twice and
 * invalidate on ItemChanged; real-host acceptance remains necessary.
 */
export function createOutlookSafetyController({ office, onState = () => {}, onInvalidated = () => {} }) {
  let ready = false, epoch = 0, revision = 0, current = null, draft = null, approval = null, inFlight = false, error = '';
  let sequence = 0;
  const identities = new WeakMap();
  const objectToken = item => { if (!identities.has(item)) identities.set(item, ++sequence); return identities.get(item); };
  const freeze = value => {
    if (value && typeof value === 'object') { Object.values(value).forEach(freeze); Object.freeze(value); }
    return value;
  };
  const itemNow = () => office?.context?.mailbox?.item || null;
  const modeOf = item => typeof item?.subject?.getAsync === 'function' ? 'compose' : 'read';
  const idOf = item => typeof item?.itemId === 'string' ? item.itemId : '';
  const readField = (field, required = false) => new Promise((resolve, reject) => {
    if (field == null) { if (required) reject(new Error('Required Outlook field unavailable.')); else resolve(''); return; }
    if (typeof field.getAsync !== 'function') { resolve(field); return; }
    try {
      field.getAsync(result => result?.status === office.AsyncResultStatus.Succeeded
        ? resolve(result.value) : reject(new Error(result?.error?.message || 'Outlook context read failed.')));
    } catch (e) { reject(e); }
  });
  const address = value => typeof value === 'string' ? value.trim().toLowerCase() : String(value?.emailAddress || '').trim().toLowerCase();
  const addresses = values => (Array.isArray(values) ? values : values ? [values] : []).map(address).filter(Boolean).sort();
  const signature = snapshot => JSON.stringify({ mode: snapshot.mode, itemId: snapshot.itemId, unsavedToken: snapshot.itemId ? null : snapshot.objectToken,
    subject: snapshot.subject, from: snapshot.from, body: snapshot.body, to: snapshot.to, cc: snapshot.cc, bcc: snapshot.bcc });
  const sameTarget = (a, b) => !!a && !!b && signature(a) === signature(b);
  const canApprove = () => ready && !inFlight && !!draft && !draft.used && sameTarget(current, draft.context);
  const getState = () => freeze({ ready, epoch, current, draft, approval, inFlight, canApprove: canApprove(), error });
  const publish = () => onState(getState());
  const invalidate = reason => {
    epoch += 1; current = null; approval = null; error = '';
    onInvalidated(reason); publish();
  };
  function assertStillSelected(item, expectedEpoch) {
    const now = itemNow();
    if (!ready || expectedEpoch !== epoch || !now || modeOf(now) !== modeOf(item)) throw new Error('Outlook selection changed.');
    if (item !== now) throw new Error('Outlook changed the current item object during verification. Refresh its context.');
    // Unsaved compose items have no stable itemId. Fail closed if their Office
    // object changes, even if a host could be returning a fresh proxy for A.
    if (!idOf(item) || !idOf(now)) {
      if (item !== now || idOf(item) !== idOf(now)) throw new Error('Unsaved Outlook item identity is uncertain.');
    } else if (idOf(item) !== idOf(now)) throw new Error('Outlook item identity changed.');
  }
  async function readSnapshot(item, expectedEpoch) {
    assertStillSelected(item, expectedEpoch);
    if (!item.body || typeof item.body.getAsync !== 'function') throw new Error('Outlook body read is unavailable.');
    const mode = modeOf(item);
    if (mode === 'read' && !idOf(item)) throw new Error('This message has no stable Outlook item ID.');
    // Compose recipient fields must be available: never replace a failed read
    // with an empty recipient set and treat that as verified.
    if (mode === 'compose' && (!item.to || !item.cc || !item.bcc)) throw new Error('Compose recipient verification is unavailable.');
    const [subject, from, to, cc, bcc, body] = await Promise.all([
      readField(item.subject, true), readField(item.from), readField(item.to, mode === 'compose'), readField(item.cc, mode === 'compose'),
      readField(item.bcc, mode === 'compose'), new Promise((resolve, reject) => {
        try { item.body.getAsync(office.CoercionType.Text, result => result?.status === office.AsyncResultStatus.Succeeded
          ? resolve(result.value) : reject(new Error(result?.error?.message || 'Outlook body read failed.'))); } catch (e) { reject(e); }
      }),
    ]);
    assertStillSelected(item, expectedEpoch);
    if (typeof subject !== 'string' || typeof body !== 'string' || (mode === 'compose' && ![to, cc, bcc].every(Array.isArray)))
      throw new Error('Outlook returned incomplete message context or recipients.');
    return freeze({ epoch: expectedEpoch, objectToken: objectToken(item), mode, itemId: idOf(item), conversationId: String(item.conversationId || ''),
      subject, from: address(from), to: addresses(to), cc: addresses(cc), bcc: addresses(bcc), body });
  }
  async function refresh() {
    invalidate('Outlook context is being refreshed.');
    const token = epoch, item = itemNow();
    if (!ready || !item) { error = 'Select an Outlook email.'; publish(); return null; }
    try {
      const snapshot = await readSnapshot(item, token);
      if (token !== epoch) return null;
      current = snapshot; error = ''; publish(); return snapshot;
    } catch (e) {
      if (token === epoch) { current = null; error = e.message; publish(); }
      return null;
    }
  }
  function captureDraftContext() {
    if (!ready || !current || inFlight) throw new Error('Wait for verified Outlook context.');
    return current;
  }
  function bindDraft(text, context = captureDraftContext()) {
    if (inFlight) throw new Error('Wait for the active Outlook insertion to finish.');
    if (typeof text !== 'string' || !text.trim()) throw new Error('Draft text is empty.');
    if (!current || context.epoch !== epoch || !sameTarget(current, context)) throw new Error('The draft target changed during generation.');
    approval = null;
    draft = freeze({ revision: ++revision, text: text.trim(), context, used: false });
    error = ''; publish(); return draft;
  }
  function beginApproval({ text, replyAll = false }) {
    if (!canApprove()) throw new Error('The draft is not bound to the verified current Outlook item.');
    if (typeof text !== 'string' || !text.trim()) throw new Error('Draft text is empty.');
    approval = freeze({ id: `approval-${epoch}-${revision}-${++sequence}`, revision: draft.revision, epoch, context: current,
      text: text.trim(), replyAll: !!replyAll });
    publish(); return approval;
  }
  function cancelApproval() { approval = null; publish(); }
  function clearDraft() { approval = null; draft = null; error = ''; publish(); }
  function write(item, approved) {
    const html = `<div>${approved.text.split(/\n\s*\n/).map(p => `<p>${p.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/\n/g, '<br>')}</p>`).join('')}</div><br>`;
    return new Promise((resolve, reject) => {
      const done = result => result?.status === office.AsyncResultStatus.Succeeded ? resolve() : reject(new Error(result?.error?.message || 'Outlook write did not confirm success.'));
      try {
        if (approved.context.mode === 'compose') {
          if (typeof item.body?.prependAsync !== 'function') throw new Error('Outlook body insertion is unavailable.');
          item.body.prependAsync(html, { coercionType: office.CoercionType.Html }, done);
        } else {
          const name = approved.replyAll ? 'displayReplyAllFormAsync' : 'displayReplyFormAsync';
          const legacy = approved.replyAll ? 'displayReplyAllForm' : 'displayReplyForm';
          if (typeof item[name] === 'function') item[name](html, done);
          else if (typeof item[legacy] === 'function') { item[legacy](html); resolve(); }
          else throw new Error('Outlook reply creation is unavailable. Open a compose reply and prepare a new draft.');
        }
      } catch (e) { reject(e); }
    });
  }
  async function confirm(approvalId) {
    if (inFlight || !approval || approval.id !== approvalId) throw new Error('That confirmation is no longer valid.');
    const approved = approval;
    // Consume before the first asynchronous read. The same click cannot replay,
    // including after a fast success callback or a failed validation.
    approval = null; inFlight = true; publish();
    try {
      if (!draft || draft.used || draft.revision !== approved.revision || approved.epoch !== epoch) throw new Error('The approval is no longer valid.');
      const item = itemNow();
      const fresh = await readSnapshot(item, approved.epoch);
      if (!sameTarget(fresh, approved.context)) throw new Error('Outlook content or recipients changed after confirmation opened.');
      if (!draft || draft.used || draft.revision !== approved.revision) throw new Error('The draft was cleared or replaced during verification.');
      assertStillSelected(item, approved.epoch);
      // Conservatively consume the draft before dispatch. If the host reports an
      // uncertain failure, require a new draft and approval instead of replaying.
      draft = freeze({ ...draft, used: true });
      await write(item, approved);
      error = ''; return { mode: approved.context.mode, text: approved.text };
    } catch (e) { error = e.message; throw e; }
    finally { inFlight = false; publish(); }
  }
  function setReady(value) { ready = !!value; invalidate(ready ? 'Outlook is ready.' : 'Outlook is unavailable.'); }
  return Object.freeze({ setReady, refresh, getState, captureDraftContext, bindDraft, beginApproval, confirm, cancelApproval, clearDraft });
}
