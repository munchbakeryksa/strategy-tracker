// Posts newly overdue tasks to the Teams chat, once per task. Run by GitHub Actions.
const { SUPABASE_URL, SUPABASE_SECRET_KEY, TEAMS_WEBHOOK_URL, GITHUB_REPOSITORY } = process.env;
for (const [k, v] of Object.entries({ SUPABASE_URL, SUPABASE_SECRET_KEY, TEAMS_WEBHOOK_URL })) {
  if (!v) { console.error(`Missing repository secret: ${k}`); process.exit(1); }
}

const base = SUPABASE_URL.trim().replace(/\/+$/, '') + '/rest/v1/items';
const key = SUPABASE_SECRET_KEY.trim();
// New secret keys (sb_secret_...) go on the apikey header only; legacy JWT keys also need Authorization.
const headers = { apikey: key, 'Content-Type': 'application/json' };
if (key.startsWith('eyJ')) headers.Authorization = `Bearer ${key}`;

// Today's date in Jeddah (UTC+3), as YYYY-MM-DD.
const today = new Intl.DateTimeFormat('en-CA', { timeZone: 'Asia/Riyadh' }).format(new Date());

const res = await fetch(
  `${base}?select=id,category,title,owner,deadline&status=neq.Done&deadline=lt.${today}&alerted_on=is.null&order=category.asc,deadline.asc`,
  { headers }
);
if (!res.ok) { console.error(`Database read failed: ${res.status} ${await res.text()}`); process.exit(1); }
const items = await res.json();
if (!items.length) { console.log('No newly overdue tasks. Nothing sent.'); process.exit(0); }

const days = d => Math.round((Date.parse(today) - Date.parse(d)) / 864e5);
const fmt = d => new Date(d + 'T00:00:00Z').toLocaleDateString('en-GB', { day: 'numeric', month: 'short', timeZone: 'UTC' });
const [owner, repo] = (GITHUB_REPOSITORY || '').split('/');
const trackerUrl = owner && repo ? `https://${owner.toLowerCase()}.github.io/${repo}/` : '';

const card = {
  $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
  type: 'AdaptiveCard',
  version: '1.4',
  body: [
    { type: 'TextBlock', size: 'Medium', weight: 'Bolder', wrap: true,
      text: `Overdue: ${items.length} task${items.length === 1 ? '' : 's'}` },
    ...items.map(i => {
      const n = days(i.deadline);
      return { type: 'TextBlock', wrap: true, spacing: 'Small',
        text: `**${i.category}** · ${i.title || 'Untitled'} · ${i.owner || 'Unassigned'} · due ${fmt(i.deadline)} (${n} day${n === 1 ? '' : 's'} overdue)` };
    })
  ],
  ...(trackerUrl ? { actions: [{ type: 'Action.OpenUrl', title: 'Open tracker', url: trackerUrl }] } : {})
};

const post = await fetch(TEAMS_WEBHOOK_URL.trim(), {
  method: 'POST',
  headers: { 'Content-Type': 'application/json' },
  body: JSON.stringify({ type: 'message', attachments: [{ contentType: 'application/vnd.microsoft.card.adaptive', contentUrl: null, content: card }] })
});
if (!post.ok) { console.error(`Teams post failed: ${post.status} ${await post.text()}`); process.exit(1); }

// Mark as alerted so each task is announced once. Changing a task's deadline in the tracker clears this.
const ids = items.map(i => `"${i.id}"`).join(',');
const mark = await fetch(`${base}?id=in.(${encodeURIComponent(ids)})`, {
  method: 'PATCH',
  headers: { ...headers, Prefer: 'return=minimal' },
  body: JSON.stringify({ alerted_on: today })
});
if (!mark.ok) { console.error(`Sent to Teams, but marking failed: ${mark.status} ${await mark.text()}`); process.exit(1); }
console.log(`Sent ${items.length} overdue task(s) to Teams.`);
