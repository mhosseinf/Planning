from flask import Flask, jsonify
from flask_cors import CORS
from pathlib import Path
import pandas as pd
from datetime import datetime, timedelta
import logging

# Create an app, being sure to pass __name__
app = Flask(__name__)
CORS(app)

# Set up logging
logging.basicConfig(level=logging.DEBUG)

# Define the file path and other constants
file_path = Path('../resource/AMSI 6 Weekly Planning (156412017).xlsm')
sheet_name = '6 Weekly'
header_row = 8
start_row = 9
end_row = 1485
start_col = 'I'
end_col = 'EL'
FieldTeam = ['Dams', 'Projects', 'CATCHMENT & ENVIRONMENT', 'CATCHMENT & ENVIRONMENT-Bunbury', 'Workshop', 'Conveyance']

# Function to process the Excel file and generate weekly and monthly plans
def generate_plans():
    logging.debug("Starting to generate plans.")
    
    try:
        data = pd.read_excel(
            file_path,
            sheet_name=sheet_name,
            header=header_row,
            usecols=f"{start_col}:{end_col}",
            skiprows=header_row,
            nrows=end_row-start_row+1
        )
        logging.debug("Excel file read successfully.")
    except Exception as e:
        logging.error(f"Error reading Excel file: {e}")
        return {}, {}

    today = datetime.today()
    current_month = today.month
    current_year = today.year
    weekly_plans = []

    # Process weekly plans
    for i in range(7):
        monday = today + timedelta(weeks=i, days=-today.weekday())
        monday = monday.date()
        logging.debug(f"Processing week starting on {monday}.")

        index_of_monday = None
        for idx, column in enumerate(data.columns):
            if isinstance(column, datetime) and column.date() == monday:
                index_of_monday = idx
                break

        if index_of_monday is not None:
            pivot_df = data.pivot_table(
                index=[data.columns[index_of_monday], 'Region.1', 'Work order', 'Proj Mgr.', 'Project Title.1', 'Task -List.1', 'Main Field Team\n/\nTeam.1'],
                values='Total Task hours.1',
                aggfunc='sum'
            )
            filtered_pivot_df = pivot_df[
                (pivot_df.index.get_level_values(data.columns[index_of_monday]) == 'x') |
                (pivot_df.index.get_level_values(data.columns[index_of_monday]) == 'X')
            ].reset_index().drop(columns=[data.columns[index_of_monday]]).rename(columns={
                'Region.1': 'Region',
                'Project Title.1': 'Project Title',
                'Task -List.1': 'Task',
                'Main Field Team\n/\nTeam.1': 'Team',
                'Total Task hours.1': 'Total hours'
            }).set_index(['Work order', 'Region', 'Proj Mgr.', 'Project Title', 'Task', 'Team'])

            weekly_plans.append(filtered_pivot_df)

    # Team-wise weekly plans
    team_weekly_plans = {team: [] for team in FieldTeam}
    for team in FieldTeam:
        for weekly_plan in weekly_plans:
            filtered_plan = weekly_plan[weekly_plan.index.get_level_values('Team') == team]
            team_weekly_plans[team].append(filtered_plan)

    # Process monthly plans
    monthly_plans = []
    for i in range(2):
        index_of_months = [idx for idx, column in enumerate(data.columns) if isinstance(column, datetime) and column.month == current_month+i and column.year == current_year]
        monthly_plans_dfs = []

        for index_of_month in index_of_months:
            month_column_name = data.columns[index_of_month]
            pivot_df = data.pivot_table(
                index=[month_column_name, 'Region.1', 'Work order', 'Project Title.1', 'Task -List.1', 'Main Field Team\n/\nTeam.1', 'Proj Mgr.'],
                values='Total Task hours.1',
                aggfunc='sum'
            )

            filtered_pivot_df = pivot_df[
                (pivot_df.index.get_level_values(month_column_name) == 'x') |
                (pivot_df.index.get_level_values(month_column_name) == 'X')
            ].reset_index().drop(columns=[month_column_name]).rename(columns={
                'Region.1': 'Region',
                'Project Title.1': 'Project Title',
                'Task -List.1': 'Task',
                'Main Field Team\n/\nTeam.1': 'Team',
                'Proj Mgr.': 'Proj Mgr.',
                'Total Task hours.1': 'Total hours'
            }).set_index(['Region', 'Work order', 'Proj Mgr.', 'Project Title', 'Task', 'Team'])

            monthly_plans_dfs.append(filtered_pivot_df)

        concatenated_df = pd.concat(monthly_plans_dfs)
        grouped_df = concatenated_df.groupby(['Work order', 'Region', 'Proj Mgr.', 'Project Title', 'Task', 'Team'])['Total hours'].sum().reset_index()
        monthly_plans.append(grouped_df)

    # Team-wise monthly plans
    team_monthly_plans = {team: [] for team in FieldTeam}
    for team in FieldTeam:
        for monthly_plan in monthly_plans:
            filtered_plan = monthly_plan.reset_index()
            filtered_plan = filtered_plan[filtered_plan['Team'] == team]
            team_monthly_plans[team].append(filtered_plan)

    logging.debug("Finished generating plans.")
    return team_weekly_plans, team_monthly_plans

#################################################
# Flask Routes
#################################################
@app.route("/api/v1.0/team_weekly_plans")
def get_team_weekly_plans():
    logging.debug("Received request for team weekly plans.")
    team_weekly_plans, _ = generate_plans()
    # Convert DataFrame to JSON serializable format
    team_weekly_plans_json = {team: [df.reset_index().to_dict(orient='records') for df in plans] for team, plans in team_weekly_plans.items()}
    return jsonify(team_weekly_plans_json)

@app.route("/api/v1.0/team_monthly_plans")
def get_team_monthly_plans():
    logging.debug("Received request for team monthly plans.")
    _, team_monthly_plans = generate_plans()
    # Convert DataFrame to JSON serializable format
    team_monthly_plans_json = {team: [df.reset_index().to_dict(orient='records') for df in plans] for team, plans in team_monthly_plans.items()}
    return jsonify(team_monthly_plans_json)

@app.route("/")
def get_routes():
    routes_dict = {
        "/api/v1.0/": "Return a JSON list of routes.",
        "/api/v1.0/team_weekly_plans": "Get team-wise weekly plans.",
        "/api/v1.0/team_monthly_plans": "Get team-wise monthly plans."
    }
    logging.debug("Received request for routes.")
    return jsonify(routes_dict)

if __name__ == "__main__":
    app.run(debug=True)