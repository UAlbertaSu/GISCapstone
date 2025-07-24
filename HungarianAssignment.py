# Sample data

origins = ['O1', 'O2', 'O3', 'O4']
destinations = ['D1', 'D2']

# Capacity of each destination
capacity = {
    'D1': 2,
    'D2': 2
}

# Travel times in minutes (origin -> destination)
travel_time = {
    'O1': {'D1': 50, 'D2': 70},
    'O2': {'D1': 55, 'D2': 40},
    'O3': {'D1': 30, 'D2': 90},
    'O4': {'D1': 65, 'D2': 50}
}

def assign_origins(origins, destinations, capacity, travel_time, max_time=60):
    assignments = {}
    remaining_capacity = capacity.copy()

    # Sort origins by number of feasible destinations (ascending)
    origins_sorted = sorted(
        origins, 
        key=lambda o: sum(travel_time[o][d] <= max_time for d in destinations)
    )

    def backtrack(i):
        if i == len(origins_sorted):
            return True  # all assigned

        o = origins_sorted[i]
        # Filter feasible destinations with capacity and time constraints
        feasible_dests = [d for d in destinations if travel_time[o][d] <= max_time and remaining_capacity[d] > 0]

        for d in feasible_dests:
            assignments[o] = d
            remaining_capacity[d] -= 1

            if backtrack(i + 1):
                return True

            # Undo assignment if no solution found down this path
            remaining_capacity[d] += 1
            del assignments[o]

        return False

    if backtrack(0):
        return assignments
    else:
        return None  # No feasible assignment

result = assign_origins(origins, destinations, capacity, travel_time)

if result:
    print("Assignments found:")
    for o, d in result.items():
        print(f"Origin {o} -> Destination {d}")
else:
    print("No feasible assignments possible under constraints.")
