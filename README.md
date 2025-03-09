# Excel To VCF Telegram Bot

This is a simple and efficient Telegram bot designed to convert Excel (.xlsx) contact files into VCF (.vcf) files that can be easily imported into your phone contacts. 📋✨

## How to Use the Bot (For Users)

- Add the bot on Telegram:
  ```bash
  @excel_convert_vcf_bot
  ```
- Start the bot by typing `/start`.
- Prepare an Excel (.xlsx) file with contact information in this format:
  - `Name` (required) or separate columns: `First Name`, `Middle Name`, `Last Name`
  - `Phone` (required)
  - `Email` (optional)
- Send the Excel file to the bot.
- The bot will convert it into a VCF file and send it back. 🎉

## How to Set Up the Code (For Developers)

- Clone the repository:
  ```bash
  https://github.com/nngeek195/Excel_to_vcf.git
  ```
- Install Python and required libraries:
  ```bash
  pip install -r requirements.txt
  ```
- Set up your Telegram bot token from [BotFather](https://t.me/BotFather).
- Add the bot token in the code:
  ```python
  BOT_TOKEN = "YOUR_TELEGRAM_BOT_TOKEN"
  ```

## How the Bot Works

- Receives an Excel file (.xlsx) from the user.
- Reads the file using Pandas.
- Combines `First Name`, `Middle Name`, and `Last Name` if they exist.
- Converts the contact information into VCF format.
- Sends the resulting .vcf file back to the user.

## Error Handling

- Checks for missing name or phone columns and notifies the user.
- Cleans up temporary files after conversion.
- Handles unexpected errors and informs the user.

## Special Notes 💎

- The bot is simple and efficient for everyday use.
- It works best with properly formatted Excel files.
- Future updates may include more advanced features.

### Happy Hacking! 💻🎉

![b676b27f-7348-4bcc-80db-f00b132aa320](https://github.com/user-attachments/assets/80005ae1-ca98-4f91-a5c2-d362e653be1c)


