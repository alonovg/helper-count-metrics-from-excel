import convertapi
import telebot

BOT_TOKEN = "Your bot token goes here"

convertapi.api_secret = 'Your api secret goes here'

bot = telebot.TeleBot(BOT_TOKEN)
