import os
import asyncio

import pandas as pd

import discord
from discord.ext import commands

# Using excel and pandas
try:
    df = pd.read_excel("ACMDataMembers.xlsx")
    credentials_dict = dict(zip(df["Name"], df["PassPhrase"]))
except Exception as e:
    print(f"Error reading Excel file: {e}")
    credentials_dict = {}


# Bot Intiation
intents = discord.Intents.all()
verifier = commands.Bot(
    command_prefix="!!",
    intents=intents
)

@verifier.event
async def on_ready():
    print("Verifier is ready to work!")

@verifier.event
async def on_member_join(member):
    HiPeoplerole = discord.utils.get(member.guild.roles, name="Hi People!")
    ACMExecutivesRole = discord.utils.get(member.guild.roles, name="Society Executive - the creeps")
    await member.add_roles(HiPeoplerole)
    await member.send("Hello, I am the official verifier of ACM and ACM-W. I am here for verification if you are a member of the society of not. If you are a part of **ACM** or **ACM-W** please enter your name to proceed with the process. If you are not a part of the society then please type \"NO\".")
    def check(m):
            return m.author == member and m.channel == member.dm_channel
    
    try:
        name_message = await verifier.wait_for('message', check=check, timeout=300.0)
        provided_name = name_message.content
        matching_names = [name for name in credentials_dict.keys() if name.lower() == provided_name.lower()]
        
        if matching_names:
            matching_name = matching_names[0]  # Use the first matching name
            await member.send(f"Welcome, {matching_name}! Please provide your password.")
            passphrase_message = await verifier.wait_for('message', check=check, timeout=300.0)
            passphrase = passphrase_message.content
            
            if credentials_dict[matching_name] == passphrase:
                await member.send("Validation successful! You are authenticated.")
                await member.add_roles(ACMExecutivesRole)
                await member.remove_roles(HiPeoplerole)
            else:
                await member.send("Validation failed! Incorrect password. Please leave and rejoin the server to restart the verification process or you can contact 'Creep Mods' from the side panel.")
        
        else:
            await member.send(f"Validation failed! {provided_name} is not recognized. This either means you're not a ACM Member or you made a typo in your name.")
    
    except asyncio.TimeoutError:
        await member.send("Oops! You seems to have timed out. You can either leave the server and rejoin to restart this process or you can contact the 'Creep Mods' in the server from the side panel.")

async def main():
    await verifier.start(os.getenv("BOT_TOKEN"))

asyncio.run(main())