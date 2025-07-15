def autopopulate_schedule(year, month, initial_schedule):
    all_days = [date(year, month, d) for d in range(1, calendar.monthrange(year, month)[1] + 1)]
    names = ["Brandon", "Tony", "Erik"]
    schedule = initial_schedule.copy()
    off_counts = Counter(schedule.values())

    def get_least_used():
        min_count = min(off_counts.values(), default=0)
        return [n for n in names if off_counts.get(n, 0) == min_count]

    # Ensure one full weekend off per person
    weekends = []
    for i in range(1, calendar.monthrange(year, month)[1] - 1):
        d = date(year, month, i)
        if d.weekday() == 4:  # Friday
            weekends.append((d, d + timedelta(days=1), d + timedelta(days=2)))

    random.shuffle(weekends)
    used_weekends = set()
    for name in names:
        for fri, sat, sun in weekends:
            if all(d not in schedule and d.month == month for d in [fri, sat, sun]):
                schedule[fri] = schedule[sat] = schedule[sun] = name
                off_counts[name] += 3
                used_weekends.update([fri, sat, sun])
                break

    # Fill remaining blanks
    for d in all_days:
        if d not in schedule:
            choices = get_least_used()
            chosen = random.choice(choices)
            schedule[d] = chosen
            off_counts[chosen] += 1

    return schedule
