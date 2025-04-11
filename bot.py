import os
import logging
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import Updater, CommandHandler, CallbackQueryHandler, MessageHandler, Filters, CallbackContext
from io import BytesIO
from docx2pdf import convert
from pdf2docx import Converter
import pythoncom

# Enable logging
logging.basicConfig(
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s', level=logging.INFO
)
logger = logging.getLogger(__name__)

# Bot token
TOKEN = "7567690696:AAH9N3X6iN9nMgiXPM1bvTbWx5oc5_4Cnk0"

# Conversion types
CONVERSION_TYPES = {
    "word_to_pdf": {"name": "Word to PDF", "input": [".docx", ".doc"], "output": ".pdf"},
    "pdf_to_word": {"name": "PDF to Word", "input": [".pdf"], "output": ".docx"},
    # Add more conversion types here
}

# Start command
def start(update: Update, context: CallbackContext) -> None:
    user = update.effective_user
    update.message.reply_text(
        f"Hi {user.first_name}! I'm a file converter bot. "
        "Send me a file or use /convert to choose conversion options."
    )

# Convert command - shows conversion options
def convert_command(update: Update, context: CallbackContext) -> None:
    keyboard = []
    
    # Create buttons for each conversion type
    for conv_type, details in CONVERSION_TYPES.items():
        keyboard.append([InlineKeyboardButton(details["name"], callback_data=conv_type)])
    
    reply_markup = InlineKeyboardMarkup(keyboard)
    update.message.reply_text("Please choose the conversion type:", reply_markup=reply_markup)

# Handle button clicks
def button(update: Update, context: CallbackContext) -> None:
    query = update.callback_query
    query.answer()
    
    # Store the selected conversion type in user data
    context.user_data["conversion_type"] = query.data
    conversion_name = CONVERSION_TYPES[query.data]["name"]
    
    query.edit_message_text(text=f"Selected: {conversion_name}. Now please send me the file to convert.")

# Handle document messages
def handle_document(update: Update, context: CallbackContext) -> None:
    if "conversion_type" not in context.user_data:
        update.message.reply_text("Please first select a conversion type using /convert")
        return
    
    conv_type = context.user_data["conversion_type"]
    input_extensions = CONVERSION_TYPES[conv_type]["input"]
    output_extension = CONVERSION_TYPES[conv_type]["output"]
    
    # Get the file
    file = update.message.document or update.message.effective_attachment
    if not file:
        update.message.reply_text("Please send a file.")
        return
    
    file_name = file.file_name
    file_extension = os.path.splitext(file_name)[1].lower()
    
    # Check if file extension is supported for this conversion
    if file_extension not in input_extensions:
        update.message.reply_text(
            f"Unsupported file type for this conversion. "
            f"Expected: {', '.join(input_extensions)}, got: {file_extension}"
        )
        return
    
    # Download the file
    file_id = file.file_id
    new_file = context.bot.get_file(file_id)
    file_bytes = BytesIO()
    new_file.download(out=file_bytes)
    file_bytes.seek(0)
    
    # Create a temporary file
    input_filename = f"input_{file_id}{file_extension}"
    output_filename = f"output_{file_id}{output_extension}"
    
    with open(input_filename, "wb") as f:
        f.write(file_bytes.getbuffer())
    
    # Perform the conversion
    try:
        update.message.reply_text("Converting your file, please wait...")
        
        if conv_type == "word_to_pdf":
            # Initialize COM for docx2pdf
            pythoncom.CoInitialize()
            convert(input_filename, output_filename)
            pythoncom.CoUninitialize()
        elif conv_type == "pdf_to_word":
            cv = Converter(input_filename)
            cv.convert(output_filename, start=0, end=None)
            cv.close()
        # Add more conversion types here
        
        # Send the converted file back
        with open(output_filename, "rb") as f:
            update.message.reply_document(
                document=f,
                filename=os.path.splitext(file_name)[0] + output_extension
            )
        
    except Exception as e:
        logger.error(f"Error during conversion: {e}")
        update.message.reply_text("Sorry, an error occurred during conversion. Please try again.")
    finally:
        # Clean up temporary files
        if os.path.exists(input_filename):
            os.remove(input_filename)
        if os.path.exists(output_filename):
            os.remove(output_filename)

# Error handler
def error(update: Update, context: CallbackContext) -> None:
    logger.warning(f'Update {update} caused error {context.error}')

def main() -> None:
    # Create the Updater and pass it your bot's token.
    updater = Updater(TOKEN)

    # Get the dispatcher to register handlers
    dispatcher = updater.dispatcher

    # Register commands
    dispatcher.add_handler(CommandHandler("start", start))
    dispatcher.add_handler(CommandHandler("convert", convert_command))
    
    # Register button handler
    dispatcher.add_handler(CallbackQueryHandler(button))
    
    # Register document handler
    dispatcher.add_handler(MessageHandler(Filters.document, handle_document))
    
    # Register error handler
    dispatcher.add_error_handler(error)

    # Start the Bot
    updater.start_polling()

    # Run the bot until you press Ctrl-C
    updater.idle()

if __name__ == '__main__':
    main()
