import os
import pandas as pd
import win32com.client as win32

# ======================================================
# SETTINGS
# ======================================================

# CSV containing recipient information
CSV_PATH = r"recipients.tsv"
PARENT_DIR = os.getcwd()
EMAIL_SUBJECT = "PLTL End-of-Semester Feedback"

EMAIL_BODY_TEMPLATE = """
Hi {first_name},

Thank you for your hard work as a PLTL leader this semester. Here is what students had to say about their PLTL experience this semester!

Best regards,
Alex Miller
(she/her/hers)
MATH 2418 PLTL Team Lead
"""

# ======================================================
# LOAD RECIPIENTS
# ======================================================

# Expected CSV format:
# FirstLast,Email

df = pd.read_csv(CSV_PATH, sep='\t')
print(df.columns)

# Connect to Outlook
outlook = win32.Dispatch('Outlook.Application')

# HELPER FUNCTION

def find_attachment(first_name, last_name):
    # Searches all MATH #### and CHEM #### folders for a matching file.

    target_base = f"{first_name}{last_name}".lower()

    for root, dirs, files in os.walk(PARENT_DIR):

        folder_name = os.path.basename(root)

        # Only search inside folders like:
        # MATH 2418
        # CHEM 2325
        if not (
            folder_name.startswith("MATH ")
            or folder_name.startswith("CHEM ")
        ):
            continue

        for file in files:

            filename_without_ext = os.path.splitext(file)[0].lower()

            if filename_without_ext == target_base:
                return os.path.join(root, file)

    return None

# ======================================================
# CREATE DRAFTS
# ======================================================

created_count = 0
missing_files = []

for index, row in df.iterrows():

    # Read values
    full_name = str(row['Name']).strip()
    email = str(row['Email']).strip()

    split_index = None

    for i in range(1, len(full_name)):
        if full_name[i].isupper():
            split_index = i
            break

    if split_index is None:
        print(f"Could not parse name: {full_name}")
        continue

    first_name = full_name[:split_index]
    last_name = full_name[split_index:]


    # Find attachment automatically
    attachment_path = find_attachment(
        first_name,
        last_name
    )

    if not attachment_path:
        missing_name = f"{first_name}{last_name}"
        missing_files.append(missing_name)
        print(f"Missing attachment: {missing_name}")
        continue

    # Create Outlook draft
    mail = outlook.CreateItem(0)

    mail.To = email
    mail.Subject = EMAIL_SUBJECT

    # Personalize body
    body = EMAIL_BODY_TEMPLATE.format(
        first_name=first_name.strip()
    )

    mail.Body = body

    # Add attachment
    mail.Attachments.Add(attachment_path)

    # Save to Drafts
    mail.Display()

    created_count += 1

    print(f"Draft created for {first_name}")

# SUMMARY
print("\n====================================")
print(f"Drafts created: {created_count}")
print(f"Missing attachments: {len(missing_files)}")

if missing_files:
    print("\nMissing:")

    for item in missing_files:
        print(f" - {item}")

print("====================================")
