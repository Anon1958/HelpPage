import win32com.client

def get_all_messages(folder):
    """Recursively collect all items (emails) from a folder and its subfolders."""
    all_items = []
    try:
        items = folder.Items
        # Convert Items to a list for easier processing
        all_items.extend(list(items))
    except Exception as e:
        print("Error reading folder:", folder.Name, e)
    
    # Recursively process subfolders
    for subfolder in folder.Folders:
        all_items.extend(get_all_messages(subfolder))
    return all_items

# Connect to Outlook
outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")

# Let's assume you want to use your primary mailbox.
# You can access it using namespace.Folders, which is a collection of mailboxes.
primary_mailbox = namespace.Folders.Item(1)  # Adjust if you have multiple mailboxes

# Collect all messages in the mailbox
all_messages = get_all_messages(primary_mailbox)

# Now, sort messages by ReceivedTime (oldest first)
# Note: ReceivedTime is a COM Date type; sorting it like a datetime works fine.
sorted_messages = sorted(all_messages, key=lambda msg: msg.ReceivedTime)

# Print out subject and received time for the first 10 emails as a test:
for msg in sorted_messages[:10]:
    print(msg.Subject, msg.ReceivedTime)
    