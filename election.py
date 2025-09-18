# Election-Project
import openpyxl
from collections import defaultdict
import re  

election_files = [
    "Elections2004.txt",
    "Elections2009.txt",
    "Elections2014.txt"
]

workbook = openpyxl.Workbook()
default_sheet = workbook.active
workbook.remove(default_sheet)

for filename in election_files:
    match = re.search(r'\d{4}', filename)
    if match:
        year = match.group()
    else:
        year = "Unknown"

    # Initialize data structures
    seats_won = defaultdict(int)
    female_contested = 0
    female_won = 0
    max_margin = 0
    highest_margin_candidate = ""

    file = open(filename, "r")
    lines = file.readlines()
    file.close()

    # Process each line
    for i in range(len(lines)):
        line = lines[i].strip()
        parts = line.split()

        # Skip invalid lines
        if len(parts) != 10:
            print("Skipping line", i + 1, "in", filename, "- Invalid format")
            continue

        # Extract data from line
        serial_no = parts[0]
        constituency = parts[1]
        winner = parts[2]
        winner_gender = parts[3].upper()
        winner_party = parts[4]
        winner_votes = int(parts[5])
        runner = parts[6]
        runner_gender = parts[7].upper()
        runner_party = parts[8]
        runner_votes = int(parts[9])

        # 1. Count seats per party
        seats_won[winner_party] += 1

        # 2. Highest winning margin
        margin = winner_votes - runner_votes
        if margin > max_margin:
            max_margin = margin
            highest_margin_candidate = winner + " (" + winner_party + ") from " + constituency

        # 3. Female candidates
        if winner_gender == "F":
            female_won += 1
            female_contested += 1
        if runner_gender == "F":
            female_contested += 1

    # Create a new sheet for this year
    sheet = workbook.create_sheet("Election " + year)

    sheet.append(["1. Seats won by each party"])
    sheet.append(["Party", "Seats Won"])
    for party in seats_won:
        sheet.append([party, seats_won[party]])

    sheet.append([])

    sheet.append(["2. Candidate with the highest winning margin"])
    sheet.append(["Candidate", "Margin"])
    sheet.append([highest_margin_candidate, max_margin])

    sheet.append([])

    sheet.append(["3. Female Candidate Statistics"])
    sheet.append(["Metric", "Count"])
    sheet.append(["Total Female Candidates Contested", female_contested])
    sheet.append(["Total Female Candidates Won", female_won])

workbook.save("election_summary.xlsx")
print("All election summaries saved to 'election_summary.xlsx'")
