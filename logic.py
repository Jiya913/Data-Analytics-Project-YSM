import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

def clean_data(df):
    df = df.copy()
    print("Missing values per column:")
    print(df.isnull().sum())

    df["Year"] = df["Year"].fillna(0)
    df["Genre"] = df["Genre"].fillna("Unknown")
    df["Rating"] = df["Rating"].fillna(0)
    df["Box Office"] = df["Box Office"].fillna(0)

    df = df.drop_duplicates()

    print(df.columns)
    df.columns = df.columns.str.strip().str.lower().str.replace(" ", "_")
    print(df.columns)

    df["year"] = df["year"].astype(int)
    df["box_office"] = df["box_office"].astype(float)
    df["title"] = df["title"].str.strip()
    df["genre"] = df["genre"].str.strip().str.title()  

    print("\nData types after cleaning:")
    print(df.dtypes)

    df.to_excel("cleaned_data.xlsx", index=False)
    print("\nCleaned data exported successfully to 'cleaned_data.xlsx'!")
    return df


def analyze_data(df):
    avg_rating = df["rating"].mean()
    total_box_office = df["box_office"].sum()
    genre_counts = df["genre"].value_counts()
    highest_rating = df["rating"].max()
    sorted_by_rating = df.sort_values(by="rating", ascending=False)

    print(f"\nAverage Rating: {avg_rating:.2f}")
    print(f"Total Box Office: {total_box_office:.2f}")
    print(f"Highest Rating: {highest_rating}")
    print("\nMovies per Genre:")
    print(genre_counts)
    print("\nMovies sorted by Rating (descending):")
    print(sorted_by_rating.head(10)) 


def visualize_data(df):
    df['genre'].value_counts().plot(kind='bar', figsize=(8,5), title="Movies per Genre")
    plt.xlabel("Genre")
    plt.ylabel("Count")
    plt.savefig("bar_chart_genre.png")
    plt.show()

    df['rating'].round().value_counts().plot(kind='pie', autopct='%1.1f%%', figsize=(6,6))
    plt.title("Rating Distribution")
    plt.ylabel("")
    plt.savefig("pie_chart_rating.png")
    plt.show()

    df.groupby('year')['box_office'].sum().plot(kind='line', marker='o', figsize=(8,5))
    plt.title("Total Box Office Collection by Year")
    plt.xlabel("Year")
    plt.ylabel("Box Office (Cr)")
    plt.savefig("line_chart_boxoffice.png")
    plt.show()

    sns.scatterplot(x='rating', y='box_office', data=df)
    plt.title("Rating vs Box Office")
    plt.xlabel("Rating")
    plt.ylabel("Box Office (Cr)")
    plt.savefig("scatter_rating_boxoffice.png")
    plt.show()

    df['rating'].plot(kind='hist', bins=10, figsize=(8,5), edgecolor='black')
    plt.title("Distribution of Ratings")
    plt.xlabel("Rating")
    plt.ylabel("Frequency")
    plt.savefig("histogram_ratings.png")
    plt.show()


def main():
    df = pd.read_excel("Movies.xlsx")
    print("Initial Data:")
    print(df)
    df = clean_data(df)
    analyze_data(df)
    visualize_data(df)

if __name__ == "__main__":
    main()

