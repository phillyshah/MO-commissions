-- Package C additions to `notification` (additive only).
--
-- `attempts`: a failed send is retried at most three times; without a counter a dead relay would
-- retry the same row forever and the row would never read "failed" to anyone looking.
-- `context`: the rendered email is built at SEND time (so the freeze token is minted when the mail
-- actually goes out), but digests carry counts computed at QUEUE time. Those counts need a home.
alter table notification add column if not exists attempts int not null default 0;
alter table notification add column if not exists context jsonb not null default '{}';
create index if not exists notification_queue_idx on notification (created_at)
  where delivery_status in ('queued', 'failed');
