# Outlook Attachment Extractor

A console application written in C# that extracts email attachments from specific folders in an Outlook PST file. Attachments are saved to a specified directory, organized by sender and subject. This tool is useful for archiving and organizing attachments from large volumes of emails.

## Features

- **Extract Attachments**: Retrieves attachments from emails in specified Outlook folders, such as "Sent Items" or custom folders.
- **Organized Storage**: Saves attachments in folders structured by sender and email subject.
- **Error Handling**: Skips attachments that have already been saved and logs any errors during the extraction process.
- **Recursive Folder Search**: Supports recursive enumeration of subfolders to extract attachments from all nested folders.
- **Customizable File Path**: The base path where attachments are saved can be easily modified.

## Requirements

- Microsoft Outlook installed on the machine.
- C# development environment (such as Visual Studio).
- .NET Framework with support for `Microsoft.Office.Interop.Outlook`.

## Installation and Setup

1. **Clone the repository**:
   ```bash
   git clone https://github.com/your-username/outlook-attachment-extractor.git
   cd outlook-attachment-extractor

2. **Open in Visual Studio**:
   - Open `OutlookAttachmentExtractor.sln` in Visual Studio.

3. **Set the Base Path**:
   - The default base path for saving attachments is `C:\temp\emails\`. You can change this in the `Program.cs` file by modifying the `basePath` variable.

4. **Add Reference to Microsoft Outlook Interop**:
   - In Visual Studio, go to **Project > Add Reference**.
   - Select **COM > Microsoft Outlook xx.x Object Library** (where `xx.x` matches your Outlook version).

## Usage

1. **Run the Application**:
   - Start the application in Visual Studio or build and run the executable.

2. **Select Folders to Process**:
   - The script defaults to processing the "Sent Items" folder in the provided PST file.
   - You can change the folder to "Inbox" or other specific folders by modifying the `EnumerateFoldersInDefaultStore` function.

3. **Process Accounts**:
   - Uncomment `EnumerateAccounts()` in `Main` to list available accounts and choose one to process specific folders for that account.

## Code Overview

The main logic is in `Program.cs`, organized into several key methods:

- **EnumerateFoldersInDefaultStore()**: Loads the PST file and initiates folder enumeration.
- **EnumerateFolders()**: Recursively navigates through subfolders, processing only the specified folders (e.g., "Sent Items").
- **IterateMessages()**: Loops through messages in each folder, extracting attachments and saving them based on sender and subject.
- **EnumerateAccountEmailAddress()**: Retrieves the email address for an account.
- **EnumerateAccounts()**: Lists all Outlook accounts for selection and processing.
- **GetFolder()**: Retrieves a specific folder based on a provided path.

### Sample Directory Structure

Attachments are saved in the following directory structure:
```
C:\temp\emails\
├── Sent Items
│   ├── [Sender Name]
│   │   └── [Subject]
│   │       ├── Attachment1.pdf
│   │       └── Attachment2.jpg
└── Inbox
    └── [Sender Name]
        └── [Subject]
            └── Attachment1.docx
```

## Example Output

```
Checking in Sent Items\John Doe
Saving: C:\temp\emails\Sent Items\John Doe\Project Proposal\Proposal.pdf
Saving: C:\temp\emails\Sent Items\John Doe\Project Proposal\Budget.xlsx
Checking in Sent Items\Jane Smith
Saving: C:\temp\emails\Sent Items\Jane Smith\Meeting Notes\Notes.docx
```

## Error Handling

If an error occurs (e.g., a folder cannot be found or an attachment cannot be saved), it is logged to the console. Common issues, such as duplicate files, are handled gracefully to prevent file overwriting.
