# In[]:
# Import the required libraries
import pandas as pd

# Import the required functions
from code_files.querying.fetch_information import fetch_schema_names, query_data
from code_files.utilities.io.database import create_sql_engine
from code_files.utilities.io.setup_db_cnxn import setup_cnxn

schema_names = fetch_schema_names()
schemas_list = list(schema_names["schema_name"])
engine = create_sql_engine()

# setup the database connection
project_root, cnxn, cursor, engine, current_dir = setup_cnxn()

# In[]:
df_breast = query_data(
    tables="[breast].[ADC_1_Breast_FollowUp_Anon] [breast].[Breast_1_FollowUp_Anon] [breast].[Post_SOC_1_Breast_FollowUp_Anon]"
)

# In[]:

# Select the desired columns only
id_cols = ["DonorID", "CohortName"]

neo_cols = df_breast.filter(regex=r"NeoadjuvantTx\d_Agent").columns.to_list()
neo_df = df_breast[id_cols + neo_cols]

adj_cols = df_breast.filter(regex=r"AdjuvantTx\d_Agent").columns.to_list()
adj_df = df_breast[id_cols + adj_cols]

subs_cols = df_breast.filter(regex=r"SubsequentTx\d_Agent").columns.to_list()
subs_df = df_breast[id_cols + subs_cols]


# In[]:

# Convert the dataframes to a long format
neo_long = pd.wide_to_long(
    neo_df,
    stubnames="NeoadjuvantTx",
    i="DonorID",
    j="Tx_Num",
    sep="",
    suffix=r"\d+_Agent",
).reset_index()
neo_long = neo_long.loc[~neo_long.NeoadjuvantTx.isna()]
neo_long[["Treatment_Type"]] = "Neoadjuvant"
neo_long = neo_long.rename(columns={"NeoadjuvantTx": "Agent"})

adj_long = pd.wide_to_long(
    adj_df, stubnames="AdjuvantTx", i="DonorID", j="Tx_Num", sep="", suffix=r"\d+_Agent"
).reset_index()
adj_long = adj_long.loc[~adj_long.AdjuvantTx.isna()]
adj_long[["Treatment_Type"]] = "Adjuvant"
adj_long = adj_long.rename(columns={"AdjuvantTx": "Agent"})

subs_long = pd.wide_to_long(
    subs_df,
    stubnames="SubsequentTx",
    i="DonorID",
    j="Tx_Num",
    sep="",
    suffix=r"\d+_Agent",
).reset_index()
# subs_long = subs_long.loc[~subs_long.SubsequentTx.isna()]
subs_long[["Treatment_Type"]] = "Subsequent"
subs_long = subs_long.rename(columns={"SubsequentTx": "Agent"})

# In[]:

# Filter each df for the ids that meet the conditions for adc treatment

adcs = [
    "Trastuzumab deruxtecan",
    "Trastuzumab emtansine",
    "Sacituzumab govitecan",
    "Datopotamab deruxtecan",
    "brentuximab",
    "Gemtuzumab ozogamicin",
    "Brentuximab vedotin",
    "Inotuzumab ozogamicin",
    "Polatuzumab vedotin",
    "Loncastuximab tesirine",
    "Pivekimab sunirine",
    "Enfortumab vedotin",
    "Mirvetuximab soravtansine",
    "Tisotumab vedotin",
    "Telisotuzumab vedotin",
]

adcs_joined = "|".join(adcs)

neo_filtered = neo_long[neo_long.Agent.str.contains(adcs_joined, case=False, na=False)]
adj_filtered = adj_long[adj_long.Agent.str.contains(adcs_joined, case=False, na=False)]
subs_filtered = subs_long[
    subs_long.Agent.str.contains(adcs_joined, case=False, na=False)
]

# In[]:

# Combine the donor ids (unique)
comb_donors = pd.concat([neo_filtered, adj_filtered, subs_filtered]).sort_values(
    by="DonorID"
)
comb_donors = comb_donors.drop_duplicates("DonorID")


#######################################################################################################################################

# In[]:
df_others = query_data(
    tables="[bladder].[Bladder_1_General_S3_FollowUp_Anon] [bladder].[Bladder_2_FollowUp_Anon] [cervical].[Cervical_1_ADC_FollowUp_Anon] [cervical].[Cervical_1_SCC_FollowUp_Anon] [longitudinal].[Longitudinal_1_Ovarian_FollowUp_Anon] [ovarian].[Ovarian_1_FollowUp_Anon] [ovarian].[Ovarian_2_S1_FollowUp_Anon] [lung].[Lung_1_S1_FollowUp_Anon] [lung].[Lung_1_S2_FollowUp_Anon] [lung].[PembroNivo_Lung_IO_1_FollowUp_Anon] [pdl1_lung].[PDL_1_lung_combined_transfer_FollowUp_Anon] [pembro_and_nivo].[Pembro_Nivo_Lung_IO_1_FollowUp_Anon] [5K_IO].[5k_IO_NSCLC_2017_2021_FollowUp_Anon] [longitudinal].[Longitudinal_1_NSCLC_FollowUp_Anon] [nsclc].[NSCLC_1_S1_FollowUp_Anon] [nsclc].[NSCLC_1_S2_FollowUp_Anon] [nsclc].[NSCLC_3_ADC_FollowUp_Anon] [nsclc].[NSCLC_3_ADC_Met_FollowUp_Anon] [nsclc].[NSCLC_3_SCC_FollowUp_Anon] [nsclc].[NSCLC_3_SCC_Met_FollowUp_Anon] [sclc].[SCLC_1_FollowUp_Anon] [sclc].[SCLC_2_FollowUp_Anon] [sclc].[SCLC_3_FollowUp_Anon] [sclc].[SCLC_4_S1_FollowUp_Anon] [sclc].[SCLC_4_S2_FollowUp_Anon]"
)

# In[]:

# Select the desired columns only
id_cols = ["DonorID", "CohortName"]

neo_cols = df_others.filter(regex=r"NeoadjuvantTx\d_Agent").columns.to_list()
neo_df = df_others[id_cols + neo_cols]
neo_df = neo_df.loc[~neo_df.NeoadjuvantTx1_Agent.isna()].drop_duplicates()

adj_cols = df_others.filter(regex=r"AdjuvantTx\d_Agent").columns.to_list()
adj_df = df_others[id_cols + adj_cols]
adj_df = adj_df.loc[~adj_df.AdjuvantTx1_Agent.isna()].drop_duplicates()

subs_cols = df_others.filter(regex=r"SubsequentTx\d_Agent").columns.to_list()
subs_df = df_others[id_cols + subs_cols]
subs_df = subs_df.loc[~subs_df.SubsequentTx1_Agent.isna()].drop_duplicates()

# In[]:

# Convert the dataframes to a long format
neo_long = pd.wide_to_long(
    neo_df,
    stubnames="NeoadjuvantTx",
    i="DonorID",
    j="Tx_Num",
    sep="",
    suffix=r"\d+_Agent",
).reset_index()
neo_long = neo_long.loc[~neo_long.NeoadjuvantTx.isna()]
neo_long[["Treatment_Type"]] = "Neoadjuvant"
neo_long = neo_long.rename(columns={"NeoadjuvantTx": "Agent"})

adj_long = pd.wide_to_long(
    adj_df, stubnames="AdjuvantTx", i="DonorID", j="Tx_Num", sep="", suffix=r"\d+_Agent"
).reset_index()
adj_long = adj_long.loc[~adj_long.AdjuvantTx.isna()]
adj_long[["Treatment_Type"]] = "Adjuvant"
adj_long = adj_long.rename(columns={"AdjuvantTx": "Agent"})

subs_long = pd.wide_to_long(
    subs_df,
    stubnames="SubsequentTx",
    i="DonorID",
    j="Tx_Num",
    sep="",
    suffix=r"\d+_Agent",
).reset_index()
subs_long = subs_long.loc[~subs_long.SubsequentTx.isna()]
subs_long[["Treatment_Type"]] = "Subsequent"
subs_long = subs_long.rename(columns={"SubsequentTx": "Agent"})

# In[]:

# Filter each df for the ids that meet the conditions for adc treatment

adcs = [
    "trastuzumab deruxtecan",
    "Trastuzumab emtansine",
    "Sacituzumab govitecan",
    "Datopotamab deruxtecan",
    "brentuximab",
    "Gemtuzumab ozogamicin",
    "Brentuximab vedotin",
    "Inotuzumab ozogamicin",
    "Polatuzumab vedotin",
    "Loncastuximab tesirine",
    "Pivekimab sunirine",
    "Enfortumab vedotin",
    "Mirvetuximab soravtansine",
    "Tisotumab vedotin",
    "Telisotuzumab vedotin",
]

adcs_joined = "|".join(adcs)

neo_filtered = neo_long[neo_long.Agent.str.contains(adcs_joined, case=False, na=False)]
adj_filtered = adj_long[adj_long.Agent.str.contains(adcs_joined, case=False, na=False)]
subs_filtered = subs_long[
    subs_long.Agent.str.contains(adcs_joined, case=False, na=False)
]

# In[]:

# Combine the donor ids (unique)
comb_donors = pd.concat([neo_filtered, adj_filtered, subs_filtered]).sort_values(
    by="DonorID"
)
# %%

#######################################################################

# In[]:

# Select the desired columns only
id_cols = ["DonorID", "CohortName"]

neo_cols = df_others.filter(regex=r"^NeoadjuvantTx\d$").columns.to_list()
neo_df = df_others[id_cols + neo_cols]
neo_df = neo_df.loc[~neo_df.NeoadjuvantTx1.isna()].drop_duplicates()

adj_cols = df_others.filter(regex=r"^AdjuvantTx\d$").columns.to_list()
adj_df = df_others[id_cols + adj_cols]
adj_df = adj_df.loc[~adj_df.AdjuvantTx1.isna()].drop_duplicates()

subs_cols = df_others.filter(regex=r"^SubsequentTx\d$").columns.to_list()
subs_df = df_others[id_cols + subs_cols]
subs_df = subs_df.loc[~subs_df.SubsequentTx1.isna()].drop_duplicates()

# In[]:

# Convert the dataframes to a long format
neo_long = pd.wide_to_long(
    neo_df, stubnames="NeoadjuvantTx", i="DonorID", j="Tx_Num", sep="", suffix=r"\d+"
).reset_index()
neo_long = neo_long.loc[~neo_long.NeoadjuvantTx.isna()]
neo_long[["Treatment_Type"]] = "Neoadjuvant"
neo_long = neo_long.rename(columns={"NeoadjuvantTx": "Tx"})

adj_long = pd.wide_to_long(
    adj_df, stubnames="AdjuvantTx", i="DonorID", j="Tx_Num", sep="", suffix=r"\d+"
).reset_index()
adj_long = adj_long.loc[~adj_long.AdjuvantTx.isna()]
adj_long[["Treatment_Type"]] = "Adjuvant"
adj_long = adj_long.rename(columns={"AdjuvantTx": "Tx"})

subs_long = pd.wide_to_long(
    subs_df, stubnames="SubsequentTx", i="DonorID", j="Tx_Num", sep="", suffix=r"\d+"
).reset_index()
subs_long = subs_long.loc[~subs_long.SubsequentTx.isna()]
subs_long[["Treatment_Type"]] = "Subsequent"
subs_long = subs_long.rename(columns={"SubsequentTx": "Tx"})

# In[]:

# Filter each df for the ids that meet the conditions for adc treatment
neo_filtered = neo_long[neo_long.Tx.str.contains("antibody", case=False, na=False)]
adj_filtered = adj_long[adj_long.Tx.str.contains("antibody", case=False, na=False)]
subs_filtered = subs_long[subs_long.Tx.str.contains("antibody", case=False, na=False)]

# In[]:

# Combine the donor ids (unique)
comb_donors = pd.concat([neo_filtered, adj_filtered, subs_filtered]).sort_values(
    by="DonorID"
)
# %%
