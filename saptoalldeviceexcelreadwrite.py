from tkinter import filedialog
import tkinter as tk
import pandas as pd
import numpy as np
import texttable

def main():
    root = tk.Tk()
    root.geometry("400x400")
    root.overrideredirect(True)

    title_bar = tk.Frame(root, bg="blue", height=30)
    title_bar.pack(fill=tk.X)

    label = tk.Label(title_bar, text="Laiko grafikų paruošimo aplikacija", fg="white", bg="blue")
    label.pack(side=tk.LEFT)

    close_button = tk.Button(title_bar, text="X", command=root.destroy, bg="red", fg="white")
    close_button.pack(side=tk.RIGHT)

    button = tk.Button(root, text="Perskaityti atsarginių detalių iš SAP .xls, .xlsx dokumentą", fg="white", bg="blue", command=browse_read_file)
    button.pack(ipadx=50, expand=True, anchor=tk.CENTER)
    button1 = tk.Button(root, text="Išeiti", fg="white", bg="blue", command=root.destroy)
    button1.pack(ipadx=92, expand=True, anchor=tk.N)

    title_bar.bind("<B1-Motion>", lambda event: move_window(event, root))
    title_bar.bind("<Button-1>", lambda event: on_title_bar_press(event, root))

    label.bind("<B1-Motion>", lambda event: move_window(event, root))
    label.bind("<Button-1>", lambda event: on_title_bar_press(event, root))

    root.mainloop()

    return "Success"


def move_window(event, root):
    root.geometry(f"+{root.winfo_pointerx() - initial_x}+{root.winfo_pointery() - initial_y}")

def on_title_bar_press(event, root):
    global initial_x, initial_y
    initial_x = event.x
    initial_y = event.y

# Sape jau yra nukastos galūnės
def browse_read_file():

    desired_tags = ['Medžiaga', 'Bendrosios atsargos', 'Medžiagos aprašas', 'Senas medž', 'Saugojimo talpykla']
    # Testinėje aplinkoje 167_Įrenginys/Linija, 168_Saugojimo talpykla produktinėje 1_Įrenginys/Linija, 2_Saugojimo talpykla 
    empty_columns = ['Code_PCE', 'Category_PCA', 'Manufacturer_PBZ', 'Minimum count_PMC', 'Booked_PRQ', 'Available_PAC', 'Cost_PCT', '168_Saugojimo talpykla', 'ID_']
    try:
        # Get the selected Excel file path
        excel_file_path = filedialog.askopenfilename(
            initialdir='/', title='Select an Excel File', filetypes=[('Excel files', '*.xls *.xlsx')]
        )

        if excel_file_path:
            # Read the Excel file into a DataFrame
            df = pd.read_excel(excel_file_path, sheet_name='RawData', usecols=desired_tags)
            df2 = df
            
            filtered_data = df.to_dict(orient='records')

            if not isinstance(filtered_data, pd.DataFrame):
                filtered_data = pd.DataFrame(filtered_data)

            filtered_data['Medžiagos aprašas'] = filtered_data['Medžiagos aprašas'].fillna('N/A')

            filtered_data.rename(columns={'Medžiaga': 'EAN_PEZ'}, inplace=True)
            filtered_data.rename(columns={'Medžiagos aprašas': 'Spare type (name)_PPN'}, inplace=True)
            filtered_data.rename(columns={'Bendrosios atsargos': 'In stock_PIS'}, inplace=True)
            filtered_data.rename(columns={'Senas medž': '167_Įrenginys/Linija'}, inplace=True)
            filtered_data.rename(columns={'Saugojimo talpykla': '168_Saugojimo talpykla'}, inplace=True)

            filtered_data[empty_columns] = np.nan
            swapped_dict = filtered_data
            swap_df_entries(swapped_dict, list(filtered_data.keys()).index('167_Įrenginys/Linija'), list(filtered_data.keys()).index('Cost_PCT'))

            filtered_data.rename(columns={'167_Įrenginys/Linija': 'torename'}, inplace=True)
            filtered_data.rename(columns={'Cost_PCT': '167_Įrenginys/Linija'}, inplace=True)

            filtered_data.rename(columns={'168_Saugojimo talpykla': 'torename2'}, inplace=True)
            filtered_data.rename(columns={'ID_': '168_Saugojimo talpykla'}, inplace=True)


            swap_df_entries(swapped_dict, list(filtered_data.keys()).index('EAN_PEZ'), list(filtered_data.keys()).index('Available_PAC'))
            filtered_data.rename(columns={'EAN_PEZ': 'torename3'}, inplace=True)
            filtered_data.rename(columns={'Available_PAC': 'EAN_PEZ'}, inplace=True)


            filtered_data.rename(columns={'Booked_PRQ': 'Cost_PCT'}, inplace=True)
            filtered_data.rename(columns={'Minimum count_PMC': 'Available_PAC'}, inplace=True)
            filtered_data.rename(columns={'Manufacturer_PBZ': 'Booked_PRQ'}, inplace=True)

            swap_df_entries(swapped_dict, list(filtered_data.keys()).index('Spare type (name)_PPN'), list(filtered_data.keys()).index('In stock_PIS'))
            filtered_data.rename(columns={'In stock_PIS': 'torename4'}, inplace=True)
            filtered_data.rename(columns={'Spare type (name)_PPN': 'In stock_PIS'}, inplace=True)


            swap_df_entries(swapped_dict, list(filtered_data.keys()).index('In stock_PIS'), list(filtered_data.keys()).index('Category_PCA'))
            filtered_data.rename(columns={'Category_PCA': 'torename5'}, inplace=True)
            filtered_data.rename(columns={'In stock_PIS': 'Category_PCA'}, inplace=True)
            filtered_data.rename(columns={'torename3': 'ID_'}, inplace=True)
            filtered_data.rename(columns={'torename4': 'Spare type (name)_PPN'}, inplace=True)
            filtered_data.rename(columns={'torename2': 'Manufacturer_PBZ'}, inplace=True)
            filtered_data.rename(columns={'Code_PCE': 'Minimum count_PMC'}, inplace=True)
            filtered_data.rename(columns={'Category_PCA': 'Code_PCE'}, inplace=True)
            filtered_data.rename(columns={'torename': 'Category_PCA'}, inplace=True)
            filtered_data.rename(columns={'torename5': 'In stock_PIS'}, inplace=True)


            filtered_data['Available_PAC'] = filtered_data['In stock_PIS'].copy()

            filtered_data['168_Saugojimo talpykla'] = df2['Saugojimo talpykla'].copy()

            filtered_data['Spare type (name)_PPN'] = filtered_data['Spare type (name)_PPN'].fillna('N/A')

            filtered_data['Minimum count_PMC'] = filtered_data['Minimum count_PMC'].fillna(0)

            filtered_data['Booked_PRQ'] = filtered_data['Booked_PRQ'].fillna(0)

            filtered_data['Cost_PCT'] = filtered_data['Cost_PCT'].fillna(0)

            swapped_dict.to_excel('output.xlsx', index=False)

            

        # All Device limitas yra 3000 per įkėlimą. Paskutinių dviejų stulpelių skaičiai gali neatitikti. 168 į 2, O 167 į 1
        # Atskirai pasiimti pirmus keturis failus ir su jais esancius pdf ir tada pridėti. 

        else:
            print('No file selected.')

    except FileNotFoundError:
        print('Error: Selected file not found.')
    except pd.errors.EmptyDataError:
        print('Error: The Excel file is empty.')
    except KeyError:
        print('Error: One or more desired tags not found in the Excel data.')


    return 'Success'

def swap_df_entries(df, col1, col2):

    df.iloc[:, [col1, col2]] = df.iloc[:, [col2, col1]]
    return df

# Pradedama main() funkcija, prieš tai python nuskaito klases, funkcijas ir kintamuosius
if __name__ == "__main__":
    initial_x = 0
    initial_y = 0
    main()