

# WhatsApp Message Sender

This is a Python script that allows you to send messages to multiple recipients on WhatsApp using Selenium and PyQt5. It provides a user interface where you can enter the recipients' phone numbers and the message you want to send.

![WhatsApp Message Sender Preview](https://github.com/samuelarjasbi/Whatsapp_message_sender/blob/main/docs/wsms.JPG?raw=true)

## Prerequisites
- Python 3.x
- PyQt5
- Selenium
- openpyxl

## Installation

1. Clone the repository:

   ```shell
   git clone https://github.com/your-username/whatsapp-message-sender.git
   ```

2. Navigate to the project directory:

   ```shell
   cd whatsapp-message-sender
   ```

3. Install the required dependencies:

   ```shell
   pip install -r requirements.txt
   ```

## Usage

1. Run the script:

   ```shell
   python main.py
   ```

2. The WhatsApp Message Sender window will appear.

3. Enter the path to the Excel file containing the list of recipients' phone numbers, or click the "Browse" button to select the file using the file dialog.

4. Enter the message you want to send in the provided text field.

5. Click the "Start Sending" button to start the message sending process.

6. A Chrome browser window will open with WhatsApp Web. Scan the QR code to log in to your WhatsApp account.

7. Once the code is scanned, click the "Code Scanned" button to start sending messages to the recipients.

8. The script will iterate through the list of recipients and send the message to each one. The sent recipients will be displayed in the list below.

9. You can close the application after the message sending is completed.

Note: Make sure you have an active internet connection and the Chrome web browser installed.

## Customization

You can customize the appearance of the application by modifying the `dark_stylesheet` variable in the script. The stylesheet uses CSS syntax, and you can adjust the colors, fonts, and other visual elements according to your preference.

## License

This project is licensed under the [MIT License](LICENSE).
