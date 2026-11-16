import streamlit as st
import pandas as pd

st.title("Welcome to my Streamlit App")
st.header("This is a header")
st.subheader("This is a subheader")
st.write("Here is a paragraph of text")

name = st.text_input("Enter your name")
age = st.slider("Select your age", 1, 100)

if st.button("Submit"):
    st.write(f"Hello {name}, you are {age} years old!")

df = pd.DataFrame(
    {
        "Numbers": [1, 2, 3, 4, 5],
        "Squares": [1, 4, 9, 16, 25]
    }
)

st.dataframe(df)
st.line_chart(df)

col1, col2 = st.columns(2)

with col1:
    st.write("Column 1 content")

with col2:
    st.write("Column 2 content")


# ----- Building an App ----- #

st.title("Simple Data Explorer")

uploaded_file = st.file_uploader(
    "Upload an Excel file", type=["xlsx"]
)

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file, header=0, sheet_name="Prostate 2 Follow Up")
    st.write("Here's a preview of your data:")
    st.dataframe(df.head())


