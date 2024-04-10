import pandas as pd
import prince
import numpy as np
from scipy.stats import chi2_contingency

file_path = "C:/Users/piotr.janczewski/Desktop/Coaching/Olko Team/IPT_CliftonInputs.xlsx"
sheet_name = "Sheet1"
def load_data(file_path, sheet_name):
    # Load data from Excel file
    df = pd.read_excel(file_path, sheet_name=sheet_name, header=0, index_col=0)
    return df

def preprocess_data(df):
    # Convert non-numeric values to NaN
    df_numeric = df.apply(pd.to_numeric, errors='coerce')
    # Check if there are still non-numeric values
    if not df_numeric.applymap(np.isreal).all().all():
        raise ValueError("Non-numeric values are present in the DataFrame")
    return df_numeric

def calculate_pearson_residuals(df_numeric):
    try:
        observed = df_numeric.values
        chi2, p, dof, expected = chi2_contingency(observed)
        pearson_residuals = (observed - expected) / np.sqrt(expected)
        pearson_residuals_df = pd.DataFrame(pearson_residuals, index=df_numeric.index, columns=df_numeric.columns)
        return pearson_residuals_df
    except ValueError as e:
        print("Error:", e)
        return None

def correspondence_analysis(df_numeric, normalization='symmetric', top_strengths=True):
    if top_strengths:
        df_numeric.iloc[:, 0:13] = df_numeric.iloc[:, 0:13].fillna(24)  # Fill NaN with middle value (24)
    # Convert DataFrame to numpy array
    observed = 35 - df_numeric.values
    
    # Apply normalization
    if normalization == 'row_principal':
        observed = observed / observed.sum(axis=1, keepdims=True)
    elif normalization == 'column_principal':
        observed = observed / observed.sum(axis=0, keepdims=True)
    
    # Compute row and column marginals
    row_marginals = observed.sum(axis=1, keepdims=True)
    col_marginals = observed.sum(axis=0, keepdims=True)
    total = observed.sum()

    # Expected frequencies under independence assumption
    expected = np.dot(row_marginals, col_marginals) / total

    # Pearson residuals
    residuals = (observed - expected) / np.sqrt(expected)
    
    # SVD of Pearson residuals
    U, s, Vt = np.linalg.svd(residuals, full_matrices=False)
    explained_variance = np.square(s) / (total - 1)
    
    # Correspondence analysis scores for rows and columns
    row_ca_scores = U[:, :2] * np.sqrt(total - 1)
    col_ca_scores = Vt[:2, :].T * np.sqrt(total - 1)
    
    # Create DataFrames for row and column coordinates
    row_ca_df = pd.DataFrame(row_ca_scores, index=df_numeric.index, columns=['X_dimension', 'Y_dimension'])
    col_ca_df = pd.DataFrame(col_ca_scores, index=df_numeric.columns, columns=['X_dimension', 'Y_dimension'])

    return row_ca_df, col_ca_df, explained_variance

def save_data(row_ca_df, col_ca_df, explained_variance):
    # Save outputs to Excel workbook
    file_path_split = file_path.split('.')
    output_file_path = file_path_split[0] + "." + file_path_split[1] + " - corresp." + file_path_split[2]
    print(file_path_split)
    with pd.ExcelWriter(output_file_path) as writer:
        row_ca_df.to_excel(writer, sheet_name='Row_CA_Scores')
        col_ca_df.to_excel(writer, sheet_name='Column_CA_Scores')
        pd.DataFrame(explained_variance, columns=['Explained_Variance']).to_excel(writer, sheet_name='Explained_Variance')

def main():
    df = load_data(file_path, sheet_name)

    # Preprocess the data
    df_numeric = preprocess_data(df)

    # Calculate Pearson residuals
    pearson_residuals = calculate_pearson_residuals(df_numeric)
    print("Pearson Residuals:")
    print(pearson_residuals)

    # Perform correspondence analysis
    normalization_choice = input("Choose normalization (symmetric, row_principal, column_principal): ")
    top_strengths = input("Focus only on top 13 strengths? (yes/no): ").lower() == 'yes'

    row_ca_df, col_ca_df, explained_variance = correspondence_analysis(df_numeric, normalization=normalization_choice, top_strengths=top_strengths)
    print("\nCorrespondence Analysis Scores:")
    print("\nPersons:")
    print(row_ca_df)
    print("\nStrengths:")
    print(col_ca_df)
    print("\nExplained Variance:")
    print("X dimension:", explained_variance[0])
    print("Y dimension:", explained_variance[1])

    save_data(row_ca_df, col_ca_df, explained_variance)

if __name__ == "__main__":
    main()
