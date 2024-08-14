
# CPM and PERT PROGRAM USING PYTHON LANGUAGE.
import pandas as pd
import networkx as nx
from math import sqrt
from scipy.stats import norm
from tabulate import tabulate
from statistics import variance as calc_variance, stdev as calc_stdev

# Read activity data from Excel file
def read_activity_data(filename):
    df = pd.read_excel(filename)
    activities = {}
    for _, row in df.iterrows():
        activity = row['Activity'].strip()
        predecessors = str(row['Precedent']).split(',')
        predecessors = [p.strip() for p in predecessors if p.strip() and p.strip() != 'nan']
        optimistic = row['A']
        most_likely = row['M']
        pessimistic = row['B']
        
        activities[activity] = {
            'optimistic': optimistic,
            'most_likely': most_likely,
            'pessimistic': pessimistic,
            'predecessors': predecessors
        }
    return activities

# Get activity data from Excel file
filename = "activities1.xlsx"
activities = read_activity_data(filename)

# Calculate expected time (TE), variance, and standard deviation for each activity
for activity, data in activities.items():
    o, m, p = data['optimistic'], data['most_likely'], data['pessimistic']
    te = (o + 4*m + p) / 6
    variance_val = ((p - o) / 6) ** 2
    std_dev = sqrt(variance_val)
    activities[activity]['TE'] = te
    activities[activity]['variance'] = variance_val
    activities[activity]['std_dev'] = std_dev

# CPM Section
print("=== Critical Path Method (CPM) ===")

# Create a directed graph
G = nx.DiGraph()

# Add nodes and edges to the graph
for activity, data in activities.items():
    G.add_node(activity, duration=data['TE'], variance=data['variance'], std_dev=data['std_dev'])
    for predecessor in data['predecessors']:
        G.add_edge(predecessor, activity)

# Verify all nodes have the 'duration', 'variance', and 'std_dev' attributes
for node in G.nodes:
    if 'duration' not in G.nodes[node]:
        G.nodes[node]['duration'] = 0
    if 'variance' not in G.nodes[node]:
        G.nodes[node]['variance'] = 0
    if 'std_dev' not in G.nodes[node]:
        G.nodes[node]['std_dev'] = 0

# Calculate earliest start and finish times
earliest_start = {}
earliest_finish = {}
for node in nx.topological_sort(G):
    es = max([earliest_finish.get(pred, 0) for pred in G.predecessors(node)], default=0)
    ef = es + G.nodes[node]['duration']
    earliest_start[node] = es
    earliest_finish[node] = ef

# Calculate latest start and finish times
latest_finish = {}
latest_start = {}
for node in reversed(list(nx.topological_sort(G))):
    lf = min([latest_start.get(succ, earliest_finish[max(earliest_finish, key=earliest_finish.get)]) for succ in G.successors(node)], default=earliest_finish[max(earliest_finish, key=earliest_finish.get)])
    ls = lf - G.nodes[node]['duration']
    latest_finish[node] = lf
    latest_start[node] = ls

# Determine the critical path
critical_path = [node for node in G.nodes if earliest_start[node] == latest_start[node]]
critical_activities = {node: node in critical_path for node in G.nodes}

# Calculate total project duration
total_project_duration = earliest_finish[max(earliest_finish, key=earliest_finish.get)]

# Prepare table data for CPM
cpm_table_data = []
for node in G.nodes:
    successors = list(G.successors(node))
    cpm_table_data.append([
        node,
        ', '.join(activities[node]['predecessors']),
        earliest_start.get(node, 'N/A'),
        earliest_finish.get(node, 'N/A'),
        latest_start.get(node, 'N/A'),
        latest_finish.get(node, 'N/A'),
        'Yes' if critical_activities[node] else 'No'
    ])

# Print CPM results in table format
cpm_headers = ["Activity", "Predecessors", "Earliest Start (ES)", "Earliest Finish (EF)", "Latest Start (LS)", "Latest Finish (LF)", "Critical"]
print(tabulate(cpm_table_data, headers=cpm_headers, tablefmt="pretty"))

# Print CPM critical path and total project duration
print("\nCritical Path: " + ' -> '.join(critical_path))
print(f"Total Project Duration: {total_project_duration} days")

# PERT Section
print("\n=== Program Evaluation Review Technique (PERT) ===")

# Prepare table data for PERT
pert_table_data = []
for activity in activities:
    te = activities[activity]['TE']
    variance = activities[activity]['variance']
    std_dev = activities[activity]['std_dev']
    z_score = (total_project_duration - earliest_finish[activity]) / total_standard_deviation if total_standard_deviation > 0 else 0
    probability = norm.cdf(z_score) if total_standard_deviation > 0 else 0
    
    pert_table_data.append([
        activity,
        ', '.join(activities[activity]['predecessors']),
        activities[activity]['optimistic'],
        activities[activity]['most_likely'],
        activities[activity]['pessimistic'],
        f"{te:.3f}",
        f"{variance:.2f}",
        f"{std_dev:.2f}",
        f"{probability:.2%}",
        'Yes' if critical_activities[activity] else 'No'
    ])

# Print PERT results in table format
pert_headers = ["Activity", "Predecessors", "A", "M", "B", "Mean Time (TE)", "Variance", "Standard Deviation", "Probability", "Critical"]
print(tabulate(pert_table_data, headers=pert_headers, tablefmt="pretty"))

# Specific Activities Analysis
print("\n=== Specific Activities Analysis ===")

# Specific activities to analyze (replace these with the activities you want to analyze)
specific_activities = ['A', 'B', 'C', 'D']  # List your specific activities here

# Extract durations of these activities
specific_durations = [activities[activity]['TE'] for activity in specific_activities if activity in activities]

# Calculate variance and standard deviation for these specific durations
if len(specific_durations) > 1:  # Variance and standard deviation are only meaningful with more than one value
    specific_variance = calc_variance(specific_durations)
    specific_std_dev = calc_stdev(specific_durations)
else:
    specific_variance = 0
    specific_std_dev = 0

# Prepare table data for specific activities
specific_table_data = []
for activity in specific_activities:
    if activity in activities:
        te = activities[activity]['TE']
        specific_table_data.append([
            activity,
            f"{te:.3f}"
        ])

# Print results in table format for specific activities
specific_table_headers = ["Activity", "Mean Time (TE)"]
print(tabulate(specific_table_data, headers=specific_table_headers, tablefmt="pretty"))
print(f"Variance of Specific Activities: {specific_variance:.2f}")
print(f"Standard Deviation of Specific Activities: {specific_std_dev:.2f}")
