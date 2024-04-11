import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from matplotlib.offsetbox import OffsetImage, AnnotationBbox

file_path = "CliftonApp/"
input_file_name = "IPT_CliftonInputs.xlsx"
sheet_name = "Sheet1"

# Define domains and colors for different strengths
strength_dict = {
    'Achiever': {'Domain': 'Executing', 'Color': 'purple'},
    'Arranger': {'Domain': 'Executing', 'Color': 'purple'},
    'Belief': {'Domain': 'Executing', 'Color': 'purple'},
    'Consistency': {'Domain': 'Executing', 'Color': 'purple'},
    'Deliberative': {'Domain': 'Executing', 'Color': 'purple'},
    'Discipline': {'Domain': 'Executing', 'Color': 'purple'},
    'Focus': {'Domain': 'Executing', 'Color': 'purple'},
    'Responsibility': {'Domain': 'Executing', 'Color': 'purple'},
    'Restorative': {'Domain': 'Executing', 'Color': 'purple'},
    'Activator': {'Domain': 'Influencing', 'Color': 'orange'},
    'Command': {'Domain': 'Influencing', 'Color': 'orange'},
    'Communication': {'Domain': 'Influencing', 'Color': 'orange'},
    'Competition': {'Domain': 'Influencing', 'Color': 'orange'},
    'Maximizer': {'Domain': 'Influencing', 'Color': 'orange'},
    'Self-Assurance': {'Domain': 'Influencing', 'Color': 'orange'},
    'Significance': {'Domain': 'Influencing', 'Color': 'orange'},
    'Woo': {'Domain': 'Influencing', 'Color': 'orange'},
    'Adaptability': {'Domain': 'Relationship building', 'Color': 'blue'},
    'Connectedness': {'Domain': 'Relationship building', 'Color': 'blue'},
    'Developer': {'Domain': 'Relationship building', 'Color': 'blue'},
    'Empathy': {'Domain': 'Relationship building', 'Color': 'blue'},
    'Harmony': {'Domain': 'Relationship building', 'Color': 'blue'},
    'Includer': {'Domain': 'Relationship building', 'Color': 'blue'},
    'Individualization': {'Domain': 'Relationship building', 'Color': 'blue'},
    'Positivity': {'Domain': 'Relationship building', 'Color': 'blue'},
    'Relator': {'Domain': 'Relationship building', 'Color': 'blue'},
    'Analytical': {'Domain': 'Strategic thinking', 'Color': 'green'},
    'Context': {'Domain': 'Strategic thinking', 'Color': 'green'},
    'Futuristic': {'Domain': 'Strategic thinking', 'Color': 'green'},
    'Ideation': {'Domain': 'Strategic thinking', 'Color': 'green'},
    'Input': {'Domain': 'Strategic thinking', 'Color': 'green'},
    'Intellection': {'Domain': 'Strategic thinking', 'Color': 'green'},
    'Learner': {'Domain': 'Strategic thinking', 'Color': 'green'},
    'Strategic': {'Domain': 'Strategic thinking', 'Color': 'green'}
}

domain_dict = {
    'Executing': {'Color':'purple'},
    'Influencing': {'Color':'orange'},
    'Relationship building': {'Color':'blue'},
    'Strategic thinking': {'Color':'green'}
    }

def open_data(file_path):
        # Read data from Excel file
    df = pd.read_excel(file_path + input_file_name, index_col=0)
    
    # Replace ranks greater than 13 with 24
    df_inv = 35 - df.replace(to_replace=df.columns[13:], value=24)

    return df, df_inv

def create_strengths_scatterplot(df):   
    # Calculate team diver (standard deviation) and power (average) for each column (Clifton strength)
    diver = df.std()
    power = df.mean()
    
    div_min, div_avg, div_max = diver.min(), diver.mean(), diver.max()
    pow_min, pow_avg, pow_max = power.min(), power.mean(), power.max()
    
    # Determine the intersection point
    intersection_point = (div_avg, pow_avg)
    
    # Create scatterplot
    fig, ax = plt.subplots(figsize=(10, 8))
    for strength in power.index:
        ax.scatter(diver[strength], power[strength],s=100, edgecolors='none', color = strength_dict[strength]['Color'])
        ax.annotate(strength, (diver[strength], power[strength]), textcoords="offset points", xytext=(-5,-5), ha='center', fontsize=8)
    
    # Set labels for quadrants
    ax.text(div_min, pow_max, 'High Power\nHigh homogenity', fontsize=12, ha='center')
    ax.text(div_max, pow_max, 'High Power\nHigh diversity', fontsize=12, ha='center')
    ax.text(div_min, pow_min, 'Limited Power\nHigh homogenity', fontsize=12, ha='center')
    ax.text(div_max, pow_min, 'Limited Power\nHigh diversity', fontsize=12, ha='center')
    
    # Set axis titles
    ax.set_xlabel('Diversity (stdev)')
    ax.set_ylabel('Power potential (average)')
    
    # Remove tick marks and labels
    ax.set_xticks([])
    ax.set_yticks([])
    
    # Add horizontal and vertical lines without labels
    ax.axhline(pow_avg, color='black')
    ax.axvline(div_avg, color='black')

    # Save the plot
    plt.savefig(file_path + ' - SWOT - 34 top 13.png')
    
    # Show plot
    plt.show()

def create_domains_scatterplot(df):
   
    # Compute scores for each domain as the average of relevant items for each person
    for domain in set(item['Domain'] for item in strength_dict.values()):
        domain_columns = [strength for strength, value in strength_dict.items() if value['Domain'] == domain]
        df[domain] = df[domain_columns].mean(axis=1)
    
    # Calculate team diver (standard deviation) and power (average) for each domain
    diver = df[[domain for domain in set(item['Domain'] for item in strength_dict.values())]].std()
    power = df[[domain for domain in set(item['Domain'] for item in strength_dict.values())]].mean()
    
    div_min, div_avg, div_max = diver.min(), diver.mean(), diver.max()
    pow_min, pow_avg, pow_max = power.min(), power.mean(), power.max()
    
    # Determine the intersection point
    intersection_point = (div_avg, pow_avg)
    
    # Create scatterplot
    fig, ax = plt.subplots(figsize=(10, 8))
    for domain in power.index:
        ax.scatter(diver[domain], power[domain], s=100, edgecolors='none',  color = domain_dict[domain]['Color'])
        ax.annotate(domain, (diver[domain], power[domain]), textcoords="offset points", xytext=(-5, -5), ha='center', fontsize=8)
    
    # Set labels for quadrants
    ax.text(div_min, pow_max, 'High Power\nHigh homogenity', fontsize=12, ha='center')
    ax.text(div_max, pow_max, 'High Power\nHigh diversity', fontsize=12, ha='center')
    ax.text(div_min, pow_min, 'Limited Power\nHigh homogenity', fontsize=12, ha='center')
    ax.text(div_max, pow_min, 'Limited Power\nHigh diversity', fontsize=12, ha='center')
    
    # Set axis titles
    ax.set_xlabel('Diversity (stdev)')
    ax.set_ylabel('Power potential (average)')
    
    # Remove tick marks and labels
    ax.set_xticks([])
    ax.set_yticks([])
    
    # Add horizontal and vertical lines without labels
    ax.axhline(pow_avg, color='black')
    ax.axvline(div_avg, color='black')

    # Save the plot
    plt.savefig(file_path + ' - SWOT - Domains.png')
    
    # Show plot
    plt.show()


# Define the scoring rules
def score_strength(rank):
    if rank <= 5:
        return 3
    elif rank <= 10:
        return 2
    elif rank <= 13:
        return 1
    else:
        return 0

def add_yoda_picture(ax, xy):
    yoda_file_name = "yoda.jpg"
    yoda_img = plt.imread(file_path + yoda_file_name)
    imagebox = OffsetImage(yoda_img, zoom=0.05)
    ab = AnnotationBbox(imagebox, xy, frameon=False)
    ax.add_artist(ab)

# Define the function to create stacked bar charts for each person
def create_stacked_bar_charts(df):
    # Initialize a list to store the colors of each bar stack
    bar_colors = []

    # Initialize a dictionary to store the domain scores for each person
    person_domain_scores = {}

    # Iterate through each person's scores
    for index, row in df.iterrows():
        domain_scores = {}
        cumulative_score = 0

        # Iterate through each domain and compute the score
        for domain, value in domain_dict.items():
            domain_columns = [strength for strength, item in strength_dict.items() if item['Domain'] == domain]
            domain_score = sum(score_strength(row[strength]) for strength in domain_columns)
            domain_scores[domain] = domain_score
            cumulative_score += domain_score

            # Create the cumulative stacked bar chart
            plt.barh(index, domain_score, color=value['Color'], edgecolor='white', linewidth=0.5, left=cumulative_score - domain_score)

        # Store the domain scores for the current person
        person_domain_scores[index] = domain_scores

    # Initialize a dictionary to store the maximum domain score for each domain across the team
    max_domain_scores = {'Executing': 0, 'Influencing': 0, 'Relationship building': 0, 'Strategic thinking': 0}

    # Iterate through each person's domain scores to compute the maximum domain score for each domain across the team
    for domain in max_domain_scores.keys():
        max_domain_scores[domain] = max(person_domain_scores.values(), key=lambda x: x[domain])[domain]

    # Iterate through each person's domain scores
    for index, domain_scores in person_domain_scores.items():
        # Set font background color for initials if the person's domain score matches the maximum score for that domain across the team
        for domain, value in domain_dict.items():
            if domain_scores[domain] == max_domain_scores[domain]:
                color = value['Color']
                plt.text(-1, index, "    ", ha='right', fontsize=12, va='center', bbox=dict(facecolor=color, alpha=0.5))
                break  # Break the loop once the color label is assigned
        
        # If the person isn't the team master of any domain, add a Yoda picture next to their initials
        if all(domain_scores[domain] != max_score for domain, max_score in max_domain_scores.items()):
            add_yoda_picture(plt.gca(), (1, index))

    # Set labels and title
    plt.ylabel('Person')
    plt.xlabel('')
    plt.title('CliftonStrengths Stacked Bar Charts')

    # Add legend for domain colors below the chart
    handles = [plt.Rectangle((0,0),1,1, color=color['Color']) for color in domain_dict.values()]
    plt.legend(handles, domain_dict.keys(), loc='upper center', bbox_to_anchor=(0.5, -0.1), ncol=4)

    # Save the plot
    plt.savefig(file_path + 'Domains by Person.png')

    # Show plot
    plt.xticks([])
    plt.yticks(rotation=0)
    plt.tight_layout()
    plt.show()


# Usage:
df, df_inv = open_data(file_path)

create_strengths_scatterplot(df_inv)

create_domains_scatterplot(df_inv)

create_stacked_bar_charts(df)

