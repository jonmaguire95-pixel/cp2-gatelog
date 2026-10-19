-- Swipe detail for saved days. One row per person per saved report, holding
-- every swipe for that person in the rows the day was analyzed from, in time
-- order. Rows go with their report: deleting a report removes its swipe rows.
-- Policies match the reports table: signed in team members only.
create table public.report_swipes (
  report_id uuid not null references public.reports(id) on delete cascade,
  guid text not null,
  ein text,
  swipes jsonb not null,
  primary key (report_id, guid)
);

create index report_swipes_guid_idx on public.report_swipes (guid);

alter table public.report_swipes enable row level security;

create policy team_read_swipes on public.report_swipes
  for select to authenticated using ((select public.is_app_user()));

create policy team_insert_swipes on public.report_swipes
  for insert to authenticated with check ((select public.is_app_user()));

create policy team_delete_swipes on public.report_swipes
  for delete to authenticated using ((select public.is_app_user()));
