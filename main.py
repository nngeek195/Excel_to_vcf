import os
import pandas as pd
from telegram import Update, Bot
from telegram.ext import Application, CommandHandler, MessageHandler, filters, CallbackContext

BOT_TOKEN = "7951513424:AAF9EDtJu0WOEJVElihuKJxU1MSwXsrl4ug"

def excel_to_vcf(file_path, output_path):
    df = pd.read_excel(file_path, engine='openpyxl')
    
    # Combine name columns if they exist
    if all(col in df.columns for col in ['First Name', 'Middle Name', 'Last Name']):
        df['Name'] = (
            df['First Name'].fillna('') + 
            ' ' + df['Middle Name'].fillna('') + 
            ' ' + df['Last Name'].fillna('')
        ).str.strip()
    elif all(col in df.columns for col in ['First Name', 'Last Name']):
        df['Name'] = (
            df['First Name'].fillna('') + 
            ' ' + df['Last Name'].fillna('')
        ).str.strip()
    elif 'Name' not in df.columns:
        raise ValueError("No valid name columns found. Please include 'Name' or 'First Name'/'Last Name'.")

    with open(output_path, 'w') as vcf:
        for _, row in df.iterrows():
            vcf.write("BEGIN:VCARD\nVERSION:3.0\n")
            
            # Write Name
            if 'Name' in df.columns and pd.notna(row['Name']):
                vcf.write(f"FN:{row['Name']}\n")
            
            # Write Phone
            if 'Phone' in df.columns and pd.notna(row['Phone']):
                vcf.write(f"TEL;TYPE=CELL:{row['Phone']}\n")
            
            # Write Email
            if 'Email' in df.columns and pd.notna(row['Email']):
                vcf.write(f"EMAIL:{row['Email']}\n")
            
            vcf.write("END:VCARD\n")

async def start(update: Update, context: CallbackContext) -> None:
    welcome_message = (
        "Hello! 📋 I'm your Excel to VCF converter bot.\n\n"
        "Please send me an Excel file (.xlsx) with contacts in this format:\n"
        "• Name (required)\n (if you need you can add three separated columns as First Name , Middle Name , Last Name)"
        "• Phone (required)\n"
        "• Email (optional)\n\n"
        "I'll convert it to a VCF file you can import into your contacts! ✨"
    )
    await update.message.reply_text(welcome_message)

async def handle_file(update: Update, context: CallbackContext) -> None:
    # Confirm file reception
    processing_message = await update.message.reply_text(
        "Received your file! 📥 Processing..."
    )

    # Download and process file
    file = update.message.document
    file_path = f"{file.file_name}"
    vcf_path = file_path.replace('.xlsx', '.vcf')

    try:
        # Download file
        bot_file = await context.bot.get_file(file.file_id)
        await bot_file.download_to_drive(file_path)

        # Convert to VCF
        excel_to_vcf(file_path, vcf_path)

        # Send results
        with open(vcf_path, 'rb') as vcf:
            await update.message.reply_document(
                document=vcf,
                filename=f"{file.file_name.replace('.xlsx', '.vcf')}",
                caption="✅ Conversion complete!"
            )

        # Final message
        await update.message.reply_text(
            "Thank you for using the bot! 😊 Let me know if you need anything else."
        )

    except Exception as e:
        await update.message.reply_text(f"❌ Error: {str(e)}")
    finally:
        # Cleanup
        if os.path.exists(file_path):
            os.remove(file_path)
        if os.path.exists(vcf_path):
            os.remove(vcf_path)
        # Delete processing message
        await context.bot.delete_message(
            chat_id=processing_message.chat_id,
            message_id=processing_message.message_id
        )

def main():
    application = Application.builder().token(BOT_TOKEN).build()

    application.add_handler(CommandHandler("start", start))
    application.add_handler(MessageHandler(
        filters.Document.MimeType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
        handle_file
    ))

    application.run_polling()

if __name__ == "__main__":
    main()
