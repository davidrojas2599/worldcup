# /// script
# dependencies = [
#     "pandas",
#     "requests",
#     "openpyxl",
#     "python-dotenv",
# ]
# ///

import os
import json
import smtplib
import pandas as pd
from dotenv import load_load, load_dotenv
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders

# Load environment variables from a local .env file
load_dotenv()

# ==========================================
# CONFIGURATION & CONSTANTS
# ==========================================
# Pulls keys safely from your .env file
NEBIUS_API_KEY = os.getenv("NEBIUS_API_KEY")
SENDER_EMAIL = os.getenv("SENDER_EMAIL")
SENDER_PASSWORD = os.getenv("SENDER_PASSWORD") # App password if using Gmail

RECIPIENT_EMAIL = "ozomatli11@gmail.com"
MODEL_NAME = "meta-llama/Llama-3.3-70B-Instruct-fast"
DATA_FILE = "fifa_world_cup_stats.xlsx"


def parse_stats_with_nebius(raw_daily_text, api_key):
    """
    Uses your standard Nebius API key to directly call the model.
    """
    if not api_key:
        raise ValueError("Critical Error: NEBIUS_API_KEY missing from environment/.env file.")
        
    url = "https://api.studio.nebius.ai/v1/chat/completions"
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json"
    }
    
    system_prompt = (
        "You are a sports data extractor. Analyze raw input text regarding World Cup matches. "
        "Return a JSON array of objects. Each object MUST represent a single match with keys exactly matching: "
        "'Match', 'Team_A', 'Team_B', 'Score_A', 'Score_B', 'Goals_Details', "
        "'Cards_Details', 'Goalie_Saves_A', 'Goalie_Saves_B', 'Key_Stats'. "
        "Do not include markdown wrappers around the JSON output, just raw parsable text."
    )
    
    payload = {
        "model": MODEL_NAME,
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": f"Extract stats from this text:\n\n{raw_daily_text}"}
        ],
        "temperature": 0.1
    }
    
    response = requests.post(url, headers=headers, json=payload)
    response.raise_for_status()
    
    try:
        clean_content = response.json()['choices'][0]['message']['content'].strip()
        if clean_content.startswith("```json"):
            clean_content = clean_content.split("```json")[1].split("```")[0].strip()
        return json.loads(clean_content)
    except Exception as e:
        print("Failed to parse JSON response from Model:", e)
        return []


def update_win_probabilities(match_results, probabilities_file="probabilities.json"):
    if os.path.exists(probabilities_file):
        with open(probabilities_file, "r") as f:
            team_probs = json.load(f)
    else:
        default_teams = ["Argentina", "France", "Brazil", "England", "Spain", "Germany", "Portugal", "Morocco"]
        team_probs = {team: round(100.0 / len(default_teams), 2) for team in default_teams}

    for match in match_results:
        tA, tB = match.get('Team_A'), match.get('Team_B')
        if tA not in team_probs: team_probs[tA] = 3.0
        if tB not in team_probs: team_probs[tB] = 3.0
        
        scoreA = int(match.get('Score_A', 0))
        scoreB = int(match.get('Score_B', 0))
        
        probA = team_probs[tA] / 100.0
        probB = team_probs[tB] / 100.0
        
        if scoreA > scoreB:
            shift = 0.15 * (1.0 - probA)
            team_probs[tA] += shift * 100
            team_probs[tB] -= shift * 100
        elif scoreB > scoreA:
            shift = 0.15 * (1.0 - probB)
            team_probs[tB] += shift * 100
            team_probs[tA] -= shift * 100
            
        team_probs[tA] = max(0.1, min(99.0, team_probs[tA]))
        team_probs[tB] = max(0.1, min(99.0, team_probs[tB]))

    total = sum(team_probs.values())
    for team in team_probs:
        team_probs[team] = round((team_probs[team] / total) * 100, 2)
        
    with open(probabilities_file, "w") as f:
        json.dump(team_probs, f, indent=4)
        
    return team_probs


def generate_excel_report(match_data, team_probs):
    df_matches = pd.DataFrame(match_data)
    prob_list = [{"Team": team, "Win Probability (%)": val} for team, val in team_probs.items()]
    df_probs = pd.DataFrame(prob_list).sort_values(by="Win Probability (%)", ascending=False)
    
    with pd.ExcelWriter(DATA_FILE, engine='openpyxl') as writer:
        df_matches.to_excel(writer, sheet_name="Match Statistics", index=False)
        df_probs.to_excel(writer, sheet_name="Live Win Predictions", index=False)


def email_report():
    if not SENDER_EMAIL or not SENDER_PASSWORD:
        print("Skipping email stage: SENDER_EMAIL or SENDER_PASSWORD missing from .env configuration.")
        return

    msg = MIMEMultipart()
    msg['From'] = SENDER_EMAIL
    msg['To'] = RECIPIENT_EMAIL
    msg['Subject'] = "Daily FIFA World Cup Analytics & Probability Matrix Updates"
    
    body = "Hi,\n\nPlease find attached the latest daily update covering stats, card timelines, goalie parameters, and updated tournament win probabilities.\n\nBest regards,\nAutomated Bot"
    msg.attach(MIMEText(body, 'plain'))
    
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE, "rb") as attachment:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f"attachment; filename= {os.path.basename(DATA_FILE)}")
            msg.attach(part)
            
    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(SENDER_EMAIL, SENDER_PASSWORD)
        server.send_message(msg)
        server.quit()
        print("Spreadsheet successfully emailed to:", RECIPIENT_EMAIL)
    except Exception as e:
        print("Failed to dispatch update report email:", e)


# ==========================================
# RUNNER
# ==========================================
if __name__ == "__main__":
    # Example match update dump
    sample_scraped_text = """
    Match Day 3 Summary: 
    Argentina played against France. Final score was Argentina 3, France 2. 
    Messi scored at minute 23 and minute 108. Di Maria scored at minute 36 for Argentina. 
    Mbappe scored at minute 80, 81 for France. 
    Yellow cards: Otamendi (Argentina) at 45', Rabiot (France) at 55'. Red Card: Paredes (Argentina) at 90'.
    Argentinian goalie Martinez had 5 saves. French goalie Lloris managed 3 saves.
    """
    
    print("Step 1: Processing raw inputs via Llama model using key loaded from .env...")
    new_matches = parse_stats_with_nebius(sample_scraped_text, NEBIUS_API_KEY)
    
    if new_matches:
        print("Step 2: Recalibrating adaptive team championship odds...")
        current_probabilities = update_win_probabilities(new_matches)
        
        print("Step 3: Compiling records into Excel layers...")
        generate_excel_report(new_matches, current_probabilities)
        
        print("Step 4: Mailing spreadsheet output...")
        email_report()
    else:
        print("No match records discovered.")