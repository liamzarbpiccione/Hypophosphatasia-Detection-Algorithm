import pandas as pd
import sys
import os
from datetime import datetime
import glob
import tkinter as tk
from tkinter import filedialog, messagebox

try:

    # Open a GUI dialog immediately to select Excel files or a folder.
    root = tk.Tk()
    root.withdraw()
    selected_files = filedialog.askopenfilenames(
        title='Select Excel file(s)',
        filetypes=[('Excel files', '*.xlsx *.xls'), ('All files', '*.*')]
    )
    if selected_files:
        excel_files = list(selected_files)
    else:
        folder = filedialog.askdirectory(title='Select folder containing Excel files')
        if not folder:
            messagebox.showerror('Selection Required', 'No folder or files selected. Exiting.')
            sys.exit(1)
        excel_files = glob.glob(os.path.join(folder, '*.xlsx')) + glob.glob(os.path.join(folder, '*.xls'))
        if not excel_files:
            messagebox.showerror('No Excel Files', f'No Excel files found in: {folder}')
            sys.exit(1)
    root.destroy()

    excel_files = list(dict.fromkeys(excel_files))
    if not excel_files:
        messagebox.showerror('No Files Selected', 'No Excel files selected. Exiting.')
        sys.exit(1)

    # Define possible column names for ALP, Phos, and MRN
    alp_possible_names = ['ALP', 'Alkaline Phosphatase', 'Alp', 'alp']
    phos_possible_names = ['Phos', 'Phosphate', 'PO4', 'Po4', 'po4']
    mrn_possible_names = ['MRN', 'Mrn', 'mrn']

    # Function to find the actual column name from possible names
    def find_column(df, possible_names):
        for name in possible_names:
            if name in df.columns:
                return name
        return None

    def parse_int(value, default):
        try:
            return int(value)
        except (ValueError, TypeError):
            return default

    def parse_float(value, default):
        try:
            return float(value)
        except (ValueError, TypeError):
            return default

    def show_filter_settings_window():
        window = tk.Tk()
        window.title('HPP Filter Settings')
        window.geometry('760x450')
        window.resizable(False, False)

        settings = {
            'apply_mrn': tk.BooleanVar(master=window, value=True),
            'min_mrn_entries': tk.StringVar(master=window, value='3'),
            'apply_alp_upper': tk.BooleanVar(master=window, value=True),
            'alp_upper': tk.StringVar(master=window, value='30'),
            'apply_egfr': tk.BooleanVar(master=window, value=True),
            'egfr_threshold': tk.StringVar(master=window, value='60'),
            'apply_phos': tk.BooleanVar(master=window, value=True),
            'phos_threshold': tk.StringVar(master=window, value='0.8'),
            'apply_conditional_phos': tk.BooleanVar(master=window, value=True),
            'alp_conditional': tk.StringVar(master=window, value='15'),
            'conditional_phos_threshold': tk.StringVar(master=window, value='0.8'),
            'cancelled': False
        }

        def toggle_state(var, widgets):
            state = 'normal' if var.get() else 'disabled'
            for widget in widgets:
                widget.config(state=state)

        row = 0
        tk.Label(window, text='Adjust filters and press Start processing when ready.', font=('Segoe UI', 11, 'bold')).grid(row=row, column=0, columnspan=4, sticky='w', padx=10, pady=(10, 8))
        row += 1

        tk.Checkbutton(window, text='Apply MRN minimum entries filter', variable=settings['apply_mrn'], command=lambda: toggle_state(settings['apply_mrn'], [min_mrn_entry])).grid(row=row, column=0, sticky='w', padx=10, pady=4)
        tk.Label(window, text='Min entries per MRN:').grid(row=row, column=1, sticky='e')
        min_mrn_entry = tk.Entry(window, textvariable=settings['min_mrn_entries'], width=10)
        min_mrn_entry.grid(row=row, column=2, sticky='w', padx=5)
        row += 1

        tk.Checkbutton(window, text='Apply ALP upper limit filter', variable=settings['apply_alp_upper'], command=lambda: toggle_state(settings['apply_alp_upper'], [alp_upper_entry])).grid(row=row, column=0, sticky='w', padx=10, pady=4)
        tk.Label(window, text='ALP <').grid(row=row, column=1, sticky='e')
        alp_upper_entry = tk.Entry(window, textvariable=settings['alp_upper'], width=10)
        alp_upper_entry.grid(row=row, column=2, sticky='w', padx=5)
        row += 1

        tk.Checkbutton(window, text='Apply eGFR filter', variable=settings['apply_egfr'], command=lambda: toggle_state(settings['apply_egfr'], [egfr_entry])).grid(row=row, column=0, sticky='w', padx=10, pady=4)
        tk.Label(window, text='eGFR ≥').grid(row=row, column=1, sticky='e')
        egfr_entry = tk.Entry(window, textvariable=settings['egfr_threshold'], width=10)
        egfr_entry.grid(row=row, column=2, sticky='w', padx=5)
        row += 1

        tk.Checkbutton(window, text='Apply Phos filter', variable=settings['apply_phos'], command=lambda: toggle_state(settings['apply_phos'], [phos_entry])).grid(row=row, column=0, sticky='w', padx=10, pady=4)
        tk.Label(window, text='Phos ≥').grid(row=row, column=1, sticky='e')
        phos_entry = tk.Entry(window, textvariable=settings['phos_threshold'], width=10)
        phos_entry.grid(row=row, column=2, sticky='w', padx=5)
        row += 1

        tk.Checkbutton(window, text='Apply conditional Phos filter', variable=settings['apply_conditional_phos'], command=lambda: toggle_state(settings['apply_conditional_phos'], [alp_conditional_entry, conditional_phos_entry])).grid(row=row, column=0, sticky='w', padx=10, pady=4)
        tk.Label(window, text='ALP ≥').grid(row=row, column=1, sticky='e')
        alp_conditional_entry = tk.Entry(window, textvariable=settings['alp_conditional'], width=10)
        alp_conditional_entry.grid(row=row, column=2, sticky='w', padx=5)
        tk.Label(window, text='Then Phos ≥').grid(row=row, column=3, sticky='e')
        conditional_phos_entry = tk.Entry(window, textvariable=settings['conditional_phos_threshold'], width=10)
        conditional_phos_entry.grid(row=row, column=4, sticky='w', padx=5)
        row += 1

        tk.Label(window, text='Info: When conditional Phos filter is enabled, Phos threshold is only enforced for rows with ALP ≥ conditional threshold.', wraplength=730, justify='left', fg='gray20').grid(row=row, column=0, columnspan=5, sticky='w', padx=10, pady=(12, 4))
        row += 1
        tk.Label(window, text='Tip: You can adjust any filter before processing. Disable a filter by unchecking it.', wraplength=730, justify='left', fg='gray20').grid(row=row, column=0, columnspan=5, sticky='w', padx=10, pady=(0, 10))
        row += 1

        def on_start():
            settings['apply_mrn'] = settings['apply_mrn'].get()
            settings['min_mrn_entries'] = parse_int(settings['min_mrn_entries'].get(), 3)
            settings['apply_alp_upper'] = settings['apply_alp_upper'].get()
            settings['alp_upper'] = parse_float(settings['alp_upper'].get(), 30.0)
            settings['apply_egfr'] = settings['apply_egfr'].get()
            settings['egfr_threshold'] = parse_float(settings['egfr_threshold'].get(), 60.0)
            settings['apply_phos'] = settings['apply_phos'].get()
            settings['phos_threshold'] = parse_float(settings['phos_threshold'].get(), 0.8)
            settings['apply_conditional_phos'] = settings['apply_conditional_phos'].get()
            settings['alp_conditional'] = parse_float(settings['alp_conditional'].get(), 15.0)
            settings['conditional_phos_threshold'] = parse_float(settings['conditional_phos_threshold'].get(), 0.8)
            window.quit()

        def on_cancel():
            settings['cancelled'] = True
            window.quit()

        button_frame = tk.Frame(window)
        button_frame.grid(row=row, column=0, columnspan=5, pady=12)
        tk.Button(button_frame, text='Start processing', command=on_start, width=16).pack(side='left', padx=6)
        tk.Button(button_frame, text='Cancel', command=on_cancel, width=10).pack(side='left', padx=6)

        window.protocol('WM_DELETE_WINDOW', on_cancel)
        window.mainloop()
        window.destroy()

        if settings['cancelled']:
            messagebox.showinfo('Cancelled', 'Filter selection cancelled. Exiting.')
            sys.exit(1)

        return settings

    filter_settings = show_filter_settings_window()
    apply_mrn = filter_settings['apply_mrn']
    min_mrn_entries = filter_settings['min_mrn_entries'] if apply_mrn else None
    apply_alp_upper = filter_settings['apply_alp_upper']
    alp_upper = filter_settings['alp_upper'] if apply_alp_upper else None
    apply_egfr = filter_settings['apply_egfr']
    egfr_threshold = filter_settings['egfr_threshold'] if apply_egfr else None
    apply_phos = filter_settings['apply_phos']
    phos_threshold = filter_settings['phos_threshold'] if apply_phos else None
    apply_conditional_phos = filter_settings['apply_conditional_phos']
    alp_conditional = filter_settings['alp_conditional'] if apply_conditional_phos else None
    conditional_phos_threshold = filter_settings['conditional_phos_threshold'] if apply_conditional_phos else None

    log_lines = []

    def log(message=''):
        log_lines.append(message)

    def show_result_window(text):
        result_window = tk.Tk()
        result_window.title('HPP Processing Results')
        result_window.geometry('700x700')
        result_window.resizable(True, True)

        frame = tk.Frame(result_window)
        frame.pack(fill='both', expand=True, padx=10, pady=10)

        text_widget = tk.Text(frame, wrap='word')
        text_widget.insert('1.0', text)
        text_widget.config(state='disabled')
        text_widget.pack(fill='both', expand=True, side='left')

        scrollbar = tk.Scrollbar(frame, command=text_widget.yview)
        scrollbar.pack(fill='y', side='right')
        text_widget.config(yscrollcommand=scrollbar.set)

        tk.Button(result_window, text='Close', command=result_window.destroy, width=12).pack(pady=8)
        result_window.mainloop()

    # Process all Excel files
    all_results = []
    
    for input_file in excel_files:
        log(f'\n========== Processing: {os.path.basename(input_file)} ==========')
        
        # Read all sheets from the Excel file
        xls = pd.ExcelFile(input_file)
        sheet_names = xls.sheet_names
        log(f'Found {len(sheet_names)} sheet(s): {", ".join(sheet_names)}')
        
        # Process each sheet
        for sheet_name in sheet_names:
            log(f'\n--- Sheet: {sheet_name} ---')
            
            # Read the sheet
            df = pd.read_excel(input_file, sheet_name=sheet_name)
            log(f'Initial rows: {len(df)}')
            
            # Find actual column names
            mrn_column = find_column(df, mrn_possible_names)
            alp_column = find_column(df, alp_possible_names)
            phos_column = find_column(df, phos_possible_names)
            egfr_column = 'eGFR' if 'eGFR' in df.columns else None
            
            log(f'Columns found - MRN: {mrn_column}, ALP: {alp_column}, Phos: {phos_column}, eGFR: {egfr_column}')
            
            # Add timestamp and source columns
            df['ProcessedDateTime'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            df['SourceFile'] = os.path.basename(input_file)
            df['SourceSheet'] = sheet_name

            # Convert columns to appropriate types
            if mrn_column:
                df[mrn_column] = df[mrn_column].astype(str)
            if alp_column:
                df[alp_column] = pd.to_numeric(df[alp_column], errors='coerce')
            if phos_column:
                df[phos_column] = pd.to_numeric(df[phos_column], errors='coerce')
            if egfr_column:
                df[egfr_column] = pd.to_numeric(df[egfr_column], errors='coerce')

            # Only include MRN with at least min_mrn_entries entries
            if apply_mrn:
                if not mrn_column:
                    log("Column 'MRN' not found in the sheet.")
                    df = pd.DataFrame()  # Set to empty instead of exiting
                else:
                    df = df.groupby(mrn_column).filter(lambda x: len(x) >= min_mrn_entries)
                    log(f'Rows after filtering by MRNs with >= {min_mrn_entries} entries: {len(df)}')
            else:
                log('MRN filter skipped.')


            # Filter: ALP < user upper limit
            if apply_alp_upper:
                if not alp_column:
                    log("Column 'ALP' not found in the sheet.")
                    df = pd.DataFrame()  # Set to empty instead of exiting
                else:
                    df = df[df[alp_column] < alp_upper]
                    log(f'Rows after ALP < {alp_upper}: {len(df)}')
            else:
                log('ALP upper limit filter skipped.')


            # Filter: Only include rows where eGFR is missing or eGFR is greater than or equal to the user threshold
            if apply_egfr:
                if egfr_column:
                    df = df[df[egfr_column].isna() | (df[egfr_column] >= egfr_threshold)]
                    log(f'Rows after filtering for eGFR >= {egfr_threshold} or missing: {len(df)}')
                else:
                    log('eGFR column not found, skipping eGFR filter.')
            else:
                log('eGFR filter skipped.')


            # Filter: Remove entries with Phos < threshold
            if apply_phos:
                if not phos_column:
                    log("Column 'Phos' not found in the sheet.")
                    df = pd.DataFrame()  # Set to empty instead of exiting
                else:
                    df = df[df[phos_column] >= phos_threshold]
                    log(f'Rows after Phos >= {phos_threshold}: {len(df)}')
            else:
                log('Phos filter skipped.')


            # Remove entries with Phos < threshold, but only for entries with ALP >= conditional threshold
            if apply_conditional_phos:
                if not phos_column:
                    log("Column 'Phos' not found in the sheet.")
                elif not alp_column:
                    log("Column 'ALP' not found in the sheet.")
                else:
                    # Apply Phos filter only to rows where ALP >= user threshold
                    mask = (df[alp_column] >= alp_conditional)
                    df = pd.concat([
                        df[mask & (df[phos_column] >= conditional_phos_threshold)],
                        df[~mask]
                    ])
                    log(f'Rows after conditional Phos < {conditional_phos_threshold} filter: {len(df)}')
            else:
                log('Conditional Phos filter skipped.')

            # Calculate minimum ALP per MRN (to sort MRNs by their lowest ALP value), then sort by MRN (by min ALP order), then by ALP ascending
            if len(df) > 0 and mrn_column and alp_column:
                min_alp_per_mrn = df.groupby(mrn_column)[alp_column].min().reset_index()
                min_alp_per_mrn = min_alp_per_mrn.sort_values(alp_column, ascending=True)
                # Create a categorical type for MRN based on this order
                mrn_order = min_alp_per_mrn[mrn_column].tolist()
                df[mrn_column] = pd.Categorical(df[mrn_column], categories=mrn_order, ordered=True)
                # Sort by MRN (by min ALP order), then by ALP ascending
                result = df.sort_values([mrn_column, alp_column], ascending=[True, True])
                log(f'Rows after sorting: {len(result)}')
                all_results.append(result)
            elif len(df) > 0:
                all_results.append(df)
                log(f'Rows in result: {len(df)}')
            else:
                log('No rows remaining after filtering.')

    # Combine all results
    if all_results:
        combined_result = pd.concat(all_results, ignore_index=True)
        log(f'\n=== Combined Results ===')
        log(f'Total rows from all files: {len(combined_result)}')
        
        # Save to a new Excel file on Desktop
        desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop')
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        output_file = os.path.join(desktop_path, f'HPP_output_{timestamp}.xlsx')
        combined_result.to_excel(output_file, index=False)
        log(f'Filtered data saved to {output_file}')
        log(f'Final row count in output: {len(combined_result)}')
    else:
        log('No data to process after filtering.')

    show_result_window('\n'.join(log_lines))
except Exception as e:
    message = f'An error occurred:\n\n{e}'
    messagebox.showerror('Error', message)
    try:
        import traceback
        log('\n*** ERROR OCCURRED ***')
        log(f'Error message: {e}')
        log('\nFull error details:')
        log(''.join(traceback.format_exc()))
        show_result_window('\n'.join(log_lines))
    except Exception:
        pass
    sys.exit(1)
