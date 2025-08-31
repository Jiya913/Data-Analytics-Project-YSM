import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import os

# ---------------------- Output Folders ----------------------
OUTPUT_DIR = "output"
CHARTS_DIR = os.path.join(OUTPUT_DIR, "charts")
os.makedirs(CHARTS_DIR, exist_ok=True)

# ---------------------- Data Cleaning ----------------------
def clean_data(df):
    df = df.copy()
    print(df.isnull().sum())

    df["Year"] = df["Year"].fillna(0).astype(int)
    df["Genre"] = df["Genre"].fillna("Unknown").str.strip().str.title()
    df["Rating"] = df["Rating"].fillna(0).astype(float)
    df["Box Office"] = df["Box Office"].fillna(0).astype(float)
    df["Title"] = df["Title"].str.strip()

    df = df.drop_duplicates()

    return df


# ---------------------- Analysis ----------------------
def analyze_data(df):
    insights = {
        "Average Rating": round(df["Rating"].mean(), 2),
        "Total Box Office": round(df["Box Office"].sum(), 2),
        "Highest Rating": df["Rating"].max(),
        "Movies Count": len(df),
        "Movies per Genre": df["Genre"].value_counts().to_dict()
    }
    return insights


# ---------------------- Export Cleaned Data ----------------------
def export_cleaned_data(df):
    filepath = os.path.join(OUTPUT_DIR, "cleaned_movies.xlsx")
    df.to_excel(filepath, index=False)
    return filepath


# ---------------------- Generate Charts ----------------------
def generate_charts(df):
    chart_paths = {}

    plt.figure(figsize=(8, 5))
    df["Genre"].value_counts().plot(kind="bar")
    plt.title("Movies per Genre")
    plt.xlabel("Genre")
    plt.ylabel("Count")
    path = f"{CHARTS_DIR}/genre_bar.png"
    plt.savefig(path)
    plt.close()
    chart_paths["Genre Bar"] = path

    plt.figure(figsize=(6, 6))
    df["Rating"].round().value_counts().plot(kind="pie", autopct='%1.1f%%')
    plt.title("Rating Distribution")
    path = f"{CHARTS_DIR}/rating_pie.png"
    plt.savefig(path)
    plt.close()
    chart_paths["Rating Pie"] = path

    plt.figure(figsize=(8, 5))
    df.groupby("Year")["Box Office"].sum().plot(kind="line", marker="o")
    plt.title("Box Office by Year")
    plt.xlabel("Year")
    plt.ylabel("Collection (Cr)")
    path = f"{CHARTS_DIR}/boxoffice_line.png"
    plt.savefig(path)
    plt.close()
    chart_paths["Box Office Line"] = path

    plt.figure(figsize=(7, 5))
    sns.scatterplot(x="Rating", y="Box Office", data=df)
    plt.title("Rating vs Box Office")
    path = f"{CHARTS_DIR}/scatter_rating_boxoffice.png"
    plt.savefig(path)
    plt.close()
    chart_paths["Scatter"] = path

    plt.figure(figsize=(8, 5))
    df["Rating"].plot(kind="hist", bins=10, edgecolor="black")
    plt.title("Distribution of Ratings")
    plt.xlabel("Rating")
    path = f"{CHARTS_DIR}/hist_ratings.png"
    plt.savefig(path)
    plt.close()
    chart_paths["Histogram"] = path

    return chart_paths


# ---------------------- Streamlit App ----------------------
def main():
    st.set_page_config(page_title="Movie Data Analysis", layout="wide")

    st.title("Movie Data Analysis App")
    st.markdown("Upload your Excel file to clean, analyze, and visualize the data!")

    uploaded_file = st.file_uploader("Upload Movie Excel File", type=["xlsx"])
    if uploaded_file is not None:
        df = pd.read_excel(uploaded_file)
        st.subheader("Initial Data")
        st.dataframe(df.head())

        st.subheader("Data Cleaning")
        df_clean = clean_data(df)
        st.success("Data cleaned successfully!")
        st.dataframe(df_clean.head())

        st.subheader("Data Analysis")
        insights = analyze_data(df_clean)
        col1, col2, col3 = st.columns(3)
        col1.metric("Movies", insights["Movies Count"])
        col2.metric("Average Rating", insights["Average Rating"])
        col3.metric("Highest Rating", insights["Highest Rating"])
        st.write("**Movies per Genre:**", insights["Movies per Genre"])

        if st.button("Export Cleaned Data"):
            filepath = export_cleaned_data(df_clean)
            with open(filepath, "rb") as f:
                st.download_button(
                    label="Download Cleaned Excel",
                    data=f,
                    file_name="cleaned_movies.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        st.subheader("Data Visualization")
        if st.button("Generate Charts"):
            chart_paths = generate_charts(df_clean)
            st.success("Charts generated successfully!")

            for name, path in chart_paths.items():
                st.image(path, caption=name, use_column_width=True)


if __name__ == "__main__":
    main()
