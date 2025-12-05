input_file = "pdf_automation/participant_names.txt"
output_file = "pdf_automation/participant_names_cleaned.txt"

with open(input_file, "r", encoding="utf-8") as f:
    lines = f.readlines()

# Remove leading numbers and dots, strip whitespace
names = []
for line in lines:
    # Remove leading number and dot (e.g., "1. Name")
    name = line.strip()
    if name:
        # Split at first space after the dot
        parts = name.split('. ', 1)
        if len(parts) == 2:
            name = parts[1]
        names.append(name)

# Sort by length (shortest to longest)
names_sorted = sorted(names, key=len)

# Write to output file
with open(output_file, "w", encoding="utf-8") as f:
    for name in names_sorted:
        f.write(name + "\n")

print(f"Cleaned and sorted names saved to {output_file}")