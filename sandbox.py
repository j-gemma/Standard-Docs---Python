import re

# Your regex pattern
pattern = r"NIKEC\s*[/,\s]\s*(?:BY|IN|EXISTING)\s([^\n]*)"


# Example strings to test
test_strings = [
    "NIKEC / BY LJSDFHVO",
    "NIKEC / by some arbitrary text",
    "NIKEC/In Electrical\nThis item is not in the kitchen equipment contract",
    "NIKEC / Existing TO REMAIN",
]

# Iterate over strings and print matches
for test in test_strings:
    match = re.match(pattern, test, re.IGNORECASE)
    if match:
        print(f"Matched: {test}")
    else:
        print(f"No Match: {test}")