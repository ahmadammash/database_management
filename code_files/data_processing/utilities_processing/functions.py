# Define a function to clean the ID columns
def clean_ids(df, columns, regex_pattern):
    """Strip all non-alphanumeric characters from specified columns."""
    for col in columns:
        if col in df.columns:
            df[f"{col}_cleaned"] = (
                df[col].astype(str).str.replace(regex_pattern, "", regex=True)
            )
    return df
