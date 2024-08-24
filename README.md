---

# 5F Glove Sales Data Collection Bot

This is a Telegram bot designed for MFest to collect data on glove sales from various retail locations. The bot interacts with users through a series of guided questions, captures the responses, and saves the information in an Excel file for further analysis. It helps the company gather important sales data about different types of gloves sold across various locations.

## Key Features

- **User Interaction**: The bot starts a conversation with users, asking them a series of predefined questions related to glove sales.
- **Data Collection**: Users provide information about the sales of different types of gloves, which the bot records.
- **Data Storage**: The collected data is stored in a dictionary and eventually saved to an Excel file (`FFResult.xlsx`) for further analysis.
- **Admin Functions**: Admin users (`dmitrokds`, `ksb2006`) can access additional features like starting the test or retrieving statistics.

## How It Works

1. **Start Command**: Users initiate interaction with the bot by sending the `/start` command. The bot greets the user and begins the questionnaire.
2. **City and Characteristics**: The bot first asks for the city of the retail location and the characteristics of the gloves that are most important to the customer.
3. **Glove Sales Questions**: The bot sequentially presents six different types of gloves to the user, asking for the approximate sales quantity for each type.
4. **Storing Responses**: The bot stores each response in a structured format, including details like the city, glove characteristics, sales quantities, and timestamps.
5. **Saving Data**: Once all questions are answered, the bot saves the collected data to `5fResult.txt` and `FFResult.xlsx`.
6. **Admin Access**: Admin users can request the collected statistics or restart the survey.

## Technical Details

- **Telegram Bot API**: The bot uses the `telebot` library to interact with Telegram's API, sending messages, receiving responses, and managing callbacks.
- **Data Management**: User responses are stored in a dictionary, which is serialized to a text file for persistence.
- **Excel File Generation**: The bot uses `openpyxl` to generate an Excel file with the collected data, organizing it into columns for easy analysis.
- **Image Handling**: The bot sends images of gloves to the user during the questionnaire to ensure clarity on the types of gloves being discussed.

## Project Structure

- **Config.py**: Contains the Telegram bot API key (`TELEGRAM_BOT_API`).
- **Glove Images**: Images corresponding to each glove type are stored in the project directory and are sent to the user during the questionnaire.
- **5fResult.txt**: A text file where user responses are temporarily stored.
- **FFResult.xlsx**: An Excel file generated at the end of each session, containing all collected data.

---
