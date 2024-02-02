import pandas as pd
import os
import requests
import dotenv

# Load environment variables
dotenv.load_dotenv()

def load_data(file_name):
    print("Loading data from file: ", file_name)
    if not os.path.exists(file_name):
        print("File not found")
        return None
    data = pd.read_excel(file_name, header=1)
    return data

def get_predictions(data):
    try:
        for index, row in data.iterrows():
            if pd.notna(row['Parsippany']):
                text_value = row['Parsippany']
                payload = {"text": text_value}
                response = requests.post(str(os.getenv('URL')), json=payload)
                jsonResponse = response.json()
                data.at[index, 'BiLSTM Prediction'] = jsonResponse['message'] 
        timestamp = pd.Timestamp.now().strftime("%Y%m%d%H%M%S")            
        data.to_excel(f"predictions_{timestamp}.xlsx", index=False)
        return True

    except Exception as e:
        print("Error: ", e)
        return False

if __name__ == "__main__":
    file_name = "TOP_COUPONS_REPORT__12_04_2023__02_02_2024.xls"
    data = load_data(file_name)
    if data is not None:
        if get_predictions(data):
            print("Predictions saved to file")
        else:
            print("Error getting predictions")
    else:
        print("Data not loaded")
