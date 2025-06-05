import pandas

def load_excel(file_path):
    try:
        df = pandas.read_excel(file_path)
        return df
    except Exception as e:
        print(f"Error loading Excel file: {e}")
        return None
    
def preprocess_data(df):
    # remove first row
    if df is not None and not df.empty:
        new_header = ["day", "number", "matt_1_1", "matt_2_1", "pom_1", "notte_1", "matt_1_2", "matt_2_2", "pom_2", "amb", "104", "agg", "ferie"]
        df = df.iloc[1:]
        df = df.iloc[:, :13]
        df.columns = new_header
        return df, new_header
    else:
        print("DataFrame is empty or None.")
        return None, None
    
def extract_docs(df, columns):
    '''
    Extracts unique doctor names from the specified columns of the DataFrame.
    Args:
        df (pandas.DataFrame): The DataFrame containing the data.
        columns (list): List of column names to extract doctor names from.
    '''
    docs = []
    for index, row in df.iterrows():
        for col in columns:
            doc = row[col]
            if type(doc) is not str:
                continue
            doc = doc.lower().strip()
            if type(doc) is not str:
                continue
            if "/" in doc or "\\" in doc:
                if "\\" in doc:
                    doc = doc.replace("\\", "/")
                doc_l = doc.split("/")
                for d in doc_l:
                    if d not in docs:
                        docs.append(d)
            elif "*" in doc:
                doc = doc.replace("*", "")
                if doc not in docs:
                    docs.append(doc)
            else:
                if doc not in docs:
                    docs.append(doc)
    return docs

mapping = {
    "matt_1_1": "M",
    "matt_2_1": "M",
    "pom_1": "P",
    "notte_1": "N",
    "matt_1_2": "M",
    "matt_2_2": "M",
    "pom_2": "P",
    "amb": "AMB",
    "104": "104",
    "agg": "AGG",
    "ferie": "F",
}

def process_data(df, docs, columns):
    '''
    Processes the DataFrame to create a mapping of doctors to their shifts.
    Args:
        df (pandas.DataFrame): The DataFrame containing the data.
        docs (list): List of unique doctor names.
        columns (list): List of column names to process.
    '''
    # Initialize a dictionary to hold the data for each doctor
    data = {}
    for doc in docs:
        data[doc] = [""] * len(df)
    data["day"] = df["day"].tolist()
    data["number"] = df["number"].tolist()
    data["number"] = [int(num) for num in data["number"]]

    # Iterate through each row in the DataFrame
    for index, row in df.iterrows():

        # Required to handle the case where a doctor is not present in the row (RIPOSO)
        doc_set = set(docs)

        for col in columns[:7]:

            # Extract the doctor name from the current column
            doc = row[col]
            if type(doc) is not str:
                continue
            doc = doc.lower().strip()

            # Process the doctor name based on the presence of special characters
            # in the case of multiple doctors in a single cell
            if "/" in doc or "\\" in doc:
                if "\\" in doc:
                    doc = doc.replace("\\", "/")
                doc_l = doc.split("/")
                for d in doc_l:
                    doc_set.discard(d)
                    mp = mapping[col]
                    if mp in data[d][index-1]:
                        pass
                    else:
                        data[d][index-1] += mp
            # Process the case where a doctor name contains an asterisk
            # in the case of "prestazione aggiuntiva"
            elif "*" in doc:
                doc = doc.replace("*", "")
                doc_set.discard(doc)
                mp = mapping[col]
                if mp in data[doc][index-1]:
                    pass
                else:
                    if len(data[doc][index-1]) == 0:
                        data[doc][index-1] = f"{mp}*"
                    else:
                        data[doc][index-1]+=f"/{mp}*"
            # Process the case where a doctor name is a single name without special characters
            else:
                doc_set.discard(doc)
                mp = mapping[col]
                if mp in data[doc][index-1]:
                    pass
                else:
                    data[doc][index-1]+=mp
        
        for col in columns[7:]:

            # Extract the doctor name from the current column
            doc = row[col]
            if type(doc) is not str:
                continue
            doc = doc.lower().strip()

            # Process the doctor name based on the presence of special characters
            # in the case of multiple doctors in a single cell
            if "/" in doc or "\\" in doc:
                if "\\" in doc:
                    doc = doc.replace("\\", "/")
                doc_l = doc.split("/")
                for d in doc_l:
                    doc_set.discard(d)
                    mp = mapping[col]
                    if mp in data[d][index-1]:
                        pass
                    else:
                        if len(data[d][index-1]) == 0:
                            data[d][index-1] = mp
                        else:
                            data[d][index-1]+=f"/{mp}"
            # Process the case where a doctor name contains an asterisk
            # in the case of "prestazione aggiuntiva"
            elif "*" in doc:
                doc = doc.replace("*", "")
                doc_set.discard(doc)
                mp = mapping[col]
                if mp in data[doc][index-1]:
                    pass
                else:
                    if len(data[doc][index-1]) == 0:
                        data[doc][index-1] = f"{mp}*"
                    else:
                        data[doc][index-1]+=f"/{mp}*"
            # Process the case where a doctor name is a single name without special characters
            else:
                doc_set.discard(doc)
                mp = mapping[col]
                if mp in data[doc][index-1]:
                    pass
                else:
                    if len(data[doc][index-1]) == 0:
                        data[doc][index-1] = mp
                    else:
                        data[doc][index-1]+=f"/{mp}"

        # Handle the case where a doctor is not present in the row (RIPOSO)
        for doc in doc_set:
            if index-2 == -1:
                continue
            # If the previous shift was "N" (notte), set the current shift to "SN" (smonto notte)
            if data[doc][index-2] == "N":
                if len(data[doc][index-1]) == 0:
                    data[doc][index-1] = "SN"
                else:
                    data[doc][index-1]+=f"/SN"
            # Else set the current shift to "R" (riposo)
            else:
                if len(data[doc][index-1]) == 0:
                    data[doc][index-1] = "R"
                else:
                    data[doc][index-1]+=f"/R"

    return data

def main():
    liberi_professionisti = ["zomparelli", "cutar"]

    file_path = 'in_turni_apr_25.xlsx'
    df = load_excel(file_path)
    df, columns = preprocess_data(df)

    if df is not None:
        print("Data loaded successfully:")

    docs = extract_docs(df, columns[2:])
    print("Extracted doctors:", docs)

    new_data = process_data(df, docs, columns[2:])
    new_new_data = {
        " ": [""] * len(new_data["day"]),
        "number": new_data["number"],
        "day": new_data["day"]
    }
    new_new_data[" "][0] = "TURNI UOC MEDICINA GENERALE HPP,"
    docs = sorted(docs)
    for doc in docs:
        if doc in liberi_professionisti:
            continue
        new_new_data[doc] = new_data[doc]
    new_new_data[""] = [""] * len(new_data["day"])
    liberi_professionisti = sorted(liberi_professionisti)
    for doc in liberi_professionisti:
        new_new_data[doc] = new_data[doc]

    new_df = pandas.DataFrame(new_new_data)
    doctors = new_df.columns
    # Put the first letter of the doctor name as capital
    doctors = [doc.capitalize() for doc in doctors]
    # Transpose the DataFrame to have doctors as rows and days as columns
    new_df = new_df.T
    # Add the first column with the doctor names capitalized
    new_df.insert(0, "doc", doctors)

    print(new_df)

    new_df.to_excel("out_aprile2025.xlsx", header=False, index=False)
    
    

def process_turni(df):
    df, columns = preprocess_data(df)

    docs = extract_docs(df, columns[2:])

    new_data = process_data(df, docs, columns[2:])
    new_df = pandas.DataFrame(new_data)
    doctors = new_df.columns
    # Put the first letter of the doctor name as capital
    doctors = [doc.capitalize() for doc in doctors]
    # Transpose the DataFrame to have doctors as rows and days as columns
    new_df = new_df.T
    # Add the first column with the doctor names
    new_df.insert(0, "doc", doctors)
    columns = new_df.columns.tolist()
    for i in range(1, len(columns)):
        columns[i] = columns[i]+1
    new_df.columns = columns

    return new_df

if __name__ == "__main__":
    main()


'''
excel deve essere tutto a celle singole ad esempio: mattina deve essere doppio
chiedere info riguardo lo / nel caso di 2 o più prestazioni
giorni di festa non so bene come gestirli
no nomi abbreviati o serve processarli?
'''

'''
inserimento manuale:
RF -> se si arriva a 6 turni il giorno dopo è RF
    la notte vale 2 turni
    capire lo SN quanto vale
RSA
'''