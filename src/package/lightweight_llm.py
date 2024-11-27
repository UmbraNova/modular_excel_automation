import pandas as pd
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.model_selection import train_test_split
from sklearn.linear_model import LogisticRegression
from sklearn.metrics import accuracy_score

data = pd.read_excel("X:/AUR/11.2024/25.11.2024/corelatie,hs,retail,season/llm_test/labeled_products_season.xlsx")

# Make sure no missing values
data = data.dropna(subset=["Description", "Code"])

X = data['Description']  # Product descriptions
y = data['Code']  # Corresponding codes

# Convert descriptions -> numerical cu "TF-IDF"
vectorizer = TfidfVectorizer(stop_words='english', max_features=15000)
X_vectorized = vectorizer.fit_transform(X)

# Split data for training and testing
X_train, X_test, y_train, y_test = train_test_split(X_vectorized, y, test_size=0.2, random_state=42)

# Antrenat logistic regression model
model = LogisticRegression(max_iter=3000)
model.fit(X_train, y_train)

# Testat modelul pe un test set
y_pred = model.predict(X_test)
print(f"Accuracy: {accuracy_score(y_test, y_pred):.2f}")

# Load new product descriptions
new_products = pd.read_excel("X:/AUR/11.2024/25.11.2024/corelatie,hs,retail,season/llm_test/new_products_season.xlsx")
new_products = new_products.dropna(subset=["Description"])  # Remove empty rows
new_descriptions = vectorizer.transform(new_products['Description'])

# Predict codes for new products
new_products['Predicted_Code'] = model.predict(new_descriptions)

new_products.to_excel("X:/AUR/11.2024/25.11.2024/corelatie,hs,retail,season/llm_test/result/updated_products.xlsx", index=False)

print("Predicted codes saved to 'updated_products.xlsx'")
