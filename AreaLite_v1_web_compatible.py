import pandas as pd


import os


# Function to browse files through a GUI (Step 1)



# Step 1 & 2: Let the user browse files and save as fixed names
def browse_and_save_files():
    arealite_input = "AREALITE_Delta.txt"
    cli_dump_input = "CLI_DUMP.txt"
    return arealite_input, cli_dump_input


# Step 3–5: Convert AREALITE_Delta.txt to TEMP.xlsx, remove backslashes
def create_temp_excel(arealite_file):
    df = pd.read_csv(arealite_file, delimiter=",", quotechar='"')
    df.replace(r'\\', '', regex=True, inplace=True)

    # Save as TEMP.xlsx in the "Sheet1" tab
    with pd.ExcelWriter("TEMP.xlsx", engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Sheet1", index=False)
    return df


# Step 6–7: Filter by 'create', add concat column, deduplicate
def process_create_section(df):
    create_df = df[df["Modification"].str.lower() == "create"].copy()
    create_df["concat"] = create_df["CommonID"].astype(str) + create_df["Moc"].astype(str)

    # Save CreateOnly tab and UniqueCreate tab with unique 'concat' values
    with pd.ExcelWriter("TEMP.xlsx", engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        create_df.to_excel(writer, sheet_name="CreateOnly", index=False)

        unique_df = create_df.drop_duplicates(subset="concat")
        unique_df.to_excel(writer, sheet_name="UniqueCreate", index=False)

    return create_df, unique_df


# Step 8–10: Open IntTOString.xlsx, create script by replacing CommonID in template
def generate_creation_script(unique_df):
    mo_create_df = pd.read_excel("IntTOString_Para.xlsx", sheet_name="MO_Create")

    with open("AREALITE_Delta_Script.txt", "w", encoding="utf-8") as f:
        for _, row in unique_df.iterrows():
            moc = row["Moc"]
            cid = row["CommonID"]
            match = mo_create_df[mo_create_df.iloc[:, 0] == moc]
            if not match.empty:
                template = match.iloc[0, 1]
                command = template.replace("CommonID", str(cid))
                f.write(command.strip() + "\n")


# Step 11: Create tabs for 'update' modifications based on CommonID
def process_update_section(df):
    update_df = df[df["Modification"].str.lower() == "update"].copy()

    with pd.ExcelWriter("TEMP.xlsx", engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        for cid, group in update_df.groupby("CommonID"):
            safe_sheet = f"CID_{str(cid)[:25]}"  # Ensure sheet name length < 31
            group.to_excel(writer, sheet_name=safe_sheet, index=False)


# Function to sort each 'CID_' sheet by the 'MocPath' column
def sort_cid_sheets():
    # Load all sheets into a dictionary
    xls = pd.read_excel("TEMP.xlsx", sheet_name=None)  # Load all sheets at once

    # Prepare a writer to save updated sheets
    with pd.ExcelWriter("TEMP.xlsx", engine="openpyxl", mode="w") as writer:
        for sheet_name, df in xls.items():
            # If it's a "CID_" sheet and contains "MocPath", sort it
            if "CID_" in sheet_name and "MocPath" in df.columns:
                df = df.sort_values(by="MocPath")

            # Save each sheet back (sorted or not)
            df.to_excel(writer, sheet_name=sheet_name, index=False)

    print("✅ All 'CID_' sheets sorted by 'MocPath' (if available) and saved.")

# Step X: Parse CLI_DUMP.txt and write FDN_MO + PARA_VALUE to "FULL_DUMP" tab
def parse_cli_dump_to_full_dump():
    fdn_list = []
    param_list = []

    with open("CLI_DUMP.txt", "r", encoding="utf-8") as file:
        lines = file.readlines()

    i = 0
    while i < len(lines):
        line = lines[i].strip()
        if line.startswith("FDN"):
            current_fdn = line
            i += 1
            while i < len(lines) and lines[i].strip():
                param = lines[i].strip()
                fdn_list.append(current_fdn)
                param_list.append(param)
                i += 1
        i += 1  # Skip empty line or move to next block

    df_full_dump = pd.DataFrame({
        "FDN_MO": fdn_list,
        "PARA_VALUE": param_list
    })

    # Write to FULL_DUMP tab
    with pd.ExcelWriter("TEMP.xlsx", engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        df_full_dump.to_excel(writer, sheet_name="FULL_DUMP", index=False)

    print("📥 'FULL_DUMP' sheet created from CLI_DUMP.txt.")

# Step X: Split PARA_VALUE column in FULL_DUMP tab using ':' and '"' as text qualifier
def split_full_dump_PARA_VALUE():
    df = pd.read_excel("TEMP.xlsx", sheet_name="FULL_DUMP")

    # Split PARA_VALUE into 'Parameter' and 'Value' using ':' delimiter
    # Handles cases like: "parameter_name":"parameter_value"
    new_cols = df["PARA_VALUE"].str.split(":", n=1, expand=True)

    df["Parameter"] = new_cols[0].str.strip().str.strip('"')
    df["Value"] = new_cols[1].str.strip().str.strip('"') if new_cols.shape[1] > 1 else ""

    # Drop original PARA_VALUE column
    df.drop(columns=["PARA_VALUE"], inplace=True)

    # Reorder columns
    df = df[["FDN_MO", "Parameter", "Value"]]

    # Save back to FULL_DUMP tab
    with pd.ExcelWriter("TEMP.xlsx", engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        df.to_excel(writer, sheet_name="FULL_DUMP", index=False)

    print("🧹 'FULL_DUMP' cleaned: 'PARA_VALUE' split into 'Parameter' and 'Value'.")

# Step X: Add properly quoted FDN_MO column to CID_ tabs
def add_fdn_column_to_cid_tabs():
    xls = pd.read_excel("TEMP.xlsx", sheet_name=None)

    with pd.ExcelWriter("TEMP.xlsx", engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        for sheet_name, df in xls.items():
            if sheet_name.startswith("CID_") and all(col in df.columns for col in ["CommonID", "MocPath"]):
                # Build the quoted FDN_MO string
                fdn_core = (
                    "SubNetwork=ONRM_ROOT_MO,MeContext=" +
                    df["CommonID"].astype(str) +
                    ",ManagedElement=" +
                    df["CommonID"].astype(str) +
                    "," +
                    df["MocPath"].astype(str)
                )
                df["FDN_MO"] = 'FDN : "' + fdn_core + '"'

                # Save the updated sheet
                df.to_excel(writer, sheet_name=sheet_name, index=False)

    print("🧩 'FDN_MO' column added to all CID_ tabs with proper quotes.")


# Step X: Add FDN_MO_Parameter column by simply concatenating FDN_MO and Parameter
def add_fdn_mo_parameter_column():
    xls = pd.read_excel("TEMP.xlsx", sheet_name=None)

    with pd.ExcelWriter("TEMP.xlsx", engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        for sheet_name, df in xls.items():
            if (sheet_name.startswith("CID_") or sheet_name == "FULL_DUMP") and \
               all(col in df.columns for col in ["FDN_MO", "Parameter"]):

                # Simple concat without any trimming
                df["FDN_MO_Parameter"] = df["FDN_MO"].astype(str) + df["Parameter"].astype(str)

                # Save updated sheet
                df.to_excel(writer, sheet_name=sheet_name, index=False)

    print("🔁 Reverted: 'FDN_MO_Parameter' column now uses full Parameter value (no trimming).")

# Step X: Enrich FULL_DUMP with GS Value from matching FDN_MO_Parameter in CID_ tabs
def enrich_full_dump_with_gs_value():
    # Load FULL_DUMP
    full_dump = pd.read_excel("TEMP.xlsx", sheet_name="FULL_DUMP")

    # Load all CID_ sheets
    xls = pd.read_excel("TEMP.xlsx", sheet_name=None)

    # Create a lookup dictionary from all CID_ sheets
    gs_lookup = {}

    for sheet_name, df in xls.items():
        if sheet_name.startswith("CID_") and all(col in df.columns for col in ["FDN_MO_Parameter", "GS Value"]):
            for _, row in df.iterrows():
                key = str(row["FDN_MO_Parameter"]).strip()
                gs_value = str(row["GS Value"]).strip()
                if key and key not in gs_lookup:  # first match wins
                    gs_lookup[key] = gs_value

    # Match and map GS Value or mark as Not_FOUND
    full_dump["GS Value"] = full_dump["FDN_MO_Parameter"].apply(
        lambda x: gs_lookup.get(str(x).strip(), "Not_FOUND")
    )

    # Write back to FULL_DUMP tab
    with pd.ExcelWriter("TEMP.xlsx", engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        full_dump.to_excel(writer, sheet_name="FULL_DUMP", index=False)

    print("🔄 'FULL_DUMP' updated with 'GS Value' from CID_ tabs — unmatched entries marked as 'Not_FOUND'.")

# Step X: Remove rows with 'Not_FOUND' in GS Value column from FULL_DUMP
def remove_not_found_from_full_dump():
    df = pd.read_excel("TEMP.xlsx", sheet_name="FULL_DUMP")

    if "GS Value" not in df.columns:
        print("⚠️ 'GS Value' column not found in FULL_DUMP.")
        return

    original_count = len(df)
    df_filtered = df[df["GS Value"].astype(str).str.strip().str.upper() != "NOT_FOUND"]
    removed_count = original_count - len(df_filtered)

    with pd.ExcelWriter("TEMP.xlsx", engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        df_filtered.to_excel(writer, sheet_name="FULL_DUMP", index=False)

    print(f"🧹 Removed {removed_count} rows with 'Not_FOUND' from FULL_DUMP.")

import re

def convert_gs_value_to_string():
    # Load FULL_DUMP
    df_full = pd.read_excel("TEMP.xlsx", sheet_name="FULL_DUMP")

    # Load FINAL tab from IntTOString_Para.xlsx
    df_final = pd.read_excel("IntTOString_Para.xlsx", sheet_name="FINAL")

    # Build parameter -> {int: string} mapping
    param_to_map = {}
    for _, row in df_final.iterrows():
        param = str(row["PARAMETER"]).strip()
        raw = str(row["INT_String"]).strip()

        # Extract (int)string pairs using regex
        pairs = re.findall(r"\((\d+)\)([^\(]+)", raw)
        mapping = {int_val: string_val.strip() for int_val, string_val in pairs}
        param_to_map[param] = mapping

    # Replace GS Value with string based on the mapping
    def map_value(row):
        param = str(row["Parameter"]).strip()
        gs_val = str(row["GS Value"]).strip()
        mapping = param_to_map.get(param, {})
        return mapping.get(gs_val, gs_val)  # fallback to original if not matched

    df_full["GS Value"] = df_full.apply(map_value, axis=1)

    # Save updated FULL_DUMP
    with pd.ExcelWriter("TEMP.xlsx", engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        df_full.to_excel(writer, sheet_name="FULL_DUMP", index=False)

    print("🔤 'GS Value' converted using FINAL tab format: (int)string.")

def bracket_list_parameters():
    df = pd.read_excel("TEMP.xlsx", sheet_name="FULL_DUMP")

    # Filter rows where 'Parameter' contains 'list' (case-insensitive)
    mask = df["Parameter"].str.contains("list", case=False, na=False)

    # Wrap GS Value in []
    df.loc[mask, "GS Value"] = df.loc[mask, "GS Value"].apply(lambda x: f"[{x}]")

    # Save updated sheet
    with pd.ExcelWriter("TEMP.xlsx", engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        df.to_excel(writer, sheet_name="FULL_DUMP", index=False)

    print("🧾 Wrapped GS Values with brackets where 'Parameter' contains 'list'.")

import pandas as pd
import re

def simplify_non_bracketed_ranges():
    df = pd.read_excel("TEMP.xlsx", sheet_name="FULL_DUMP")

    def process_gs_value(val):
        if pd.isna(val):
            return val  # Keep NaNs as is

        val_str = str(val).strip()

        # Skip already bracketed values
        if "[" in val_str or "]" in val_str:
            return val_str

        # Check for ".." or "," in the value
        if ".." in val_str or "," in val_str:
            # Match a leading signed integer: -12, +3, or 0
            match = re.match(r"^\s*([-+]?\d+)", val_str)
            if match:
                return match.group(1)

        return val_str  # No match — leave unchanged

    df["GS Value"] = df["GS Value"].apply(process_gs_value)

    with pd.ExcelWriter("TEMP.xlsx", engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        df.to_excel(writer, sheet_name="FULL_DUMP", index=False)

    print("✅ Cleaned GS Value: Extracted signed integer from unbracketed ranges/lists.")

import pandas as pd

def update_gs_value_conditionally():
    # Load FULL_DUMP from TEMP.xlsx
    df = pd.read_excel("TEMP.xlsx", sheet_name="FULL_DUMP")

    # Load all lines from CLI_DUMP.txt
    with open("CLI_DUMP.txt", "r", encoding="utf-8") as f:
        cli_lines = [line.strip() for line in f if line.strip().startswith("FDN")]

    for idx, row in df.iterrows():
        gs_val = str(row.get("GS Value", "")).strip()
        parameter = str(row.get("Parameter", "")).strip().lower()
        fdn_mo = str(row.get("FDN_MO", "")).strip()

        # Condition: GS Value contains both '=' and ','
        if "=" in gs_val and "," in gs_val:
            if "userl" in parameter:
                df.at[idx, "GS Value"] = f'"{gs_val}"'
            else:
                # Extract prefix from FDN_MO (first two comma-separated parts)
                fdn_parts = fdn_mo.split(",")
                if len(fdn_parts) >= 2:
                    fdn_prefix = ",".join(fdn_parts[:2]).strip()

                    match_line = None
                    for line in cli_lines:
                        if fdn_prefix in line and gs_val in line:
                            match_line = line
                            break

                    if match_line:
                        # Remove only 'FDN : ' prefix
                        updated_value = match_line.replace("FDN : ", "", 1).strip()
                        df.at[idx, "GS Value"] = updated_value
                    else:
                        print(f"❌ No match found in CLI_DUMP.txt for GS='{gs_val}' and prefix='{fdn_prefix}'")
                else:
                    print(f"⚠️ Row {idx}: Invalid FDN_MO → {fdn_mo}")

    # Save updated FULL_DUMP back to TEMP.xlsx
    with pd.ExcelWriter("TEMP.xlsx", engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        df.to_excel(writer, sheet_name="FULL_DUMP", index=False)

    print("✅ GS Value updated using full matching FDN from CLI_DUMP or quoted userLabel.")

import pandas as pd

def fix_gs_value_based_on_cli_dump():
    # Load FULL_DUMP
    df = pd.read_excel("TEMP.xlsx", sheet_name="FULL_DUMP")

    # Read CLI_DUMP.txt and keep only lines starting with 'FDN'
    with open("CLI_DUMP.txt", "r", encoding="utf-8") as f:
        cli_lines = [line.strip() for line in f if line.strip().startswith("FDN")]

    for idx, row in df.iterrows():
        gs_val = str(row.get("GS Value", "")).strip()
        fdn_mo = str(row.get("FDN_MO", "")).strip()

        # Filter: GS Value contains '=' and does NOT contain '"'
        if "=" in gs_val and '"' not in gs_val:
            fdn_parts = fdn_mo.split(",")
            if len(fdn_parts) >= 2:
                fdn_prefix = ",".join(fdn_parts[:2]).strip()

                # Look for a matching FDN line
                matching_line = None
                for line in cli_lines:
                    if fdn_prefix in line and gs_val in line:
                        matching_line = line
                        break

                if matching_line:
                    # Remove 'FDN : ' only once
                    updated_value = matching_line.replace("FDN : ", "", 1).strip()
                    df.at[idx, "GS Value"] = updated_value
                else:
                    print(f"❌ No match in CLI_DUMP for GS='{gs_val}' and FDN_MO prefix='{fdn_prefix}'")
            else:
                print(f"⚠️ Invalid FDN_MO at row {idx}: {fdn_mo}")

    # Save the updated FULL_DUMP sheet
    with pd.ExcelWriter("TEMP.xlsx", engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        df.to_excel(writer, sheet_name="FULL_DUMP", index=False)

    print("✅ GS Value updated using exact CLI_DUMP match (with '=' and no quotes).")

import pandas as pd

def override_gs_value_from_special_category():
    # Load FULL_DUMP
    full_df = pd.read_excel("TEMP.xlsx", sheet_name="FULL_DUMP")

    # Load Special_Catagory tab
    special_df = pd.read_excel("IntTOString_Para.xlsx", sheet_name="Special_Catagory")

    # Create a mapping from Parameter → GS Value
    special_map = dict(zip(special_df["Parameter"].astype(str), special_df["GS Value"].astype(str)))

    # Replace GS Value in FULL_DUMP if Parameter matches
    for idx, row in full_df.iterrows():
        param = str(row.get("Parameter", "")).strip()
        if param in special_map:
            full_df.at[idx, "GS Value"] = special_map[param]

    # Write updated FULL_DUMP back to TEMP.xlsx
    with pd.ExcelWriter("TEMP.xlsx", engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        full_df.to_excel(writer, sheet_name="FULL_DUMP", index=False)

    print("✅ GS Value updated for all matching Parameters using Special_Catagory tab.")

import pandas as pd

def extract_struct_parameters_from_update():
    # Load Sheet1 from TEMP.xlsx
    df = pd.read_excel("TEMP.xlsx", sheet_name="Sheet1")

    # Filter where Modification is 'update' (case-insensitive) and Parameter contains '.'
    filtered_df = df[
        df["Modification"].astype(str).str.lower() == "update"
    ]
    filtered_df = filtered_df[
        filtered_df["Parameter"].astype(str).str.contains(r"\.", regex=True)
    ]

    # Write to new sheet 'Struct_Para'
    with pd.ExcelWriter("TEMP.xlsx", engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        filtered_df.to_excel(writer, sheet_name="Struct_Para", index=False)

    print("📄 'Struct_Para' sheet created with parameters containing '.' and Modification='update'.")

import pandas as pd

def update_full_fdn_and_sort():
    df = pd.read_excel("TEMP.xlsx", sheet_name="Struct_Para")

    # Create FULL_FDN
    df["FULL_FDN"] = (
        "SubNetwork=ONRM_ROOT_MO,MeContext=" +
        df["CommonID"].astype(str) +
        ",ManagedElement=" +
        df["CommonID"].astype(str) +
        "," +
        df["MocPath"].astype(str)
    )

    # Sort by FULL_FDN + Parameter
    df_sorted = df.sort_values(by=["FULL_FDN", "Parameter"], ignore_index=True)

    # Save back to Excel
    with pd.ExcelWriter("TEMP.xlsx", engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        df_sorted.to_excel(writer, sheet_name="Struct_Para", index=False)

    print("✅ 'FULL_FDN' created and Struct_Para sorted by 'FULL_FDN' + 'Parameter'.")

import pandas as pd

def add_struct_parameter_column():
    # Load Struct_Para
    df = pd.read_excel("TEMP.xlsx", sheet_name="Struct_Para")

    # Clean quotes and split Parameter by '.'
    param_clean = df["Parameter"].astype(str).str.strip('"')
    struct_parts = param_clean.str.split(".", n=1, expand=True)

    # Insert Struct_Parameter column right after 'Parameter'
    insert_idx = df.columns.get_loc("Parameter") + 1
    df.insert(loc=insert_idx, column="Struct_Parameter", value=struct_parts[1])

    # Save back to Struct_Para
    with pd.ExcelWriter("TEMP.xlsx", engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        df.to_excel(writer, sheet_name="Struct_Para", index=False)

    print("🧩 'Struct_Parameter' column inserted and populated from 'Parameter' split.")

import re
import pandas as pd

def resolve_struct_parameter_gs_value():
    # Load Struct_Para
    df_struct = pd.read_excel("TEMP.xlsx", sheet_name="Struct_Para")

    # Load FINAL tab from IntTOString_Para.xlsx
    df_final = pd.read_excel("IntTOString_Para.xlsx", sheet_name="FINAL")

    # Create mapping: PARAMETER → INT_String
    param_to_intstring = dict(zip(df_final["PARAMETER"].astype(str), df_final["INT_String"].astype(str)))

    for idx, row in df_struct.iterrows():
        struct_param = str(row.get("Struct_Parameter", "")).strip()
        gs_value = str(row.get("GS Value", "")).strip()

        if struct_param in param_to_intstring:
            int_string_data = param_to_intstring[struct_param]

            # Parse INT_String into a dictionary {int: label}
            matches = re.findall(r"\((\-?\d+)\)([^\(]+)", int_string_data)
            int_to_label = {k: v.strip() for k, v in matches}

            # Replace GS Value with corresponding label if match found
            if gs_value in int_to_label:
                df_struct.at[idx, "GS Value"] = int_to_label[gs_value]

    # Save back to Struct_Para
    with pd.ExcelWriter("TEMP.xlsx", engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        df_struct.to_excel(writer, sheet_name="Struct_Para", index=False)

    print("🔄 GS Value updated in 'Struct_Para' using INT_String from FINAL tab.")

import pandas as pd

def clean_parameter_and_extract_new_gs():
    # Load Struct_Para
    df_struct = pd.read_excel("TEMP.xlsx", sheet_name="Struct_Para")

    # Step 1: Clean 'Parameter' by removing everything after the first dot
    df_struct["Parameter"] = df_struct["Parameter"].astype(str).str.split(".").str[0].str.strip()

    # Step 2: Read CLI_DUMP.txt
    with open("CLI_DUMP.txt", "r", encoding="utf-8") as f:
        lines = [line.strip() for line in f.readlines()]

    # Step 3: Organize CLI dump into FDN blocks
    blocks = []
    current_block = []
    for line in lines:
        if line.startswith("FDN :"):
            if current_block:
                blocks.append(current_block)
            current_block = [line]
        elif current_block:
            current_block.append(line)
    if current_block:
        blocks.append(current_block)

    # Step 4: Map FDN (stripped of 'FDN : ') to their parameter blocks
    fdn_block_map = {}
    for block in blocks:
        fdn_line = block[0].replace("FDN : ", "").strip().strip('"')
        fdn_block_map[fdn_line] = block[1:]

    # Step 5: Search for matching parameter line in block
    new_gs_list = []
    for _, row in df_struct.iterrows():
        full_fdn = str(row.get("FULL_FDN", "")).strip()
        param = str(row.get("Parameter", "")).strip()

        new_gs_value = ""
        if full_fdn in fdn_block_map:
            block_lines = fdn_block_map[full_fdn]
            for line in block_lines:
                if param in line:
                    new_gs_value = line
                    break

        new_gs_list.append(new_gs_value)

    # Step 6: Save results
    df_struct["New_GS"] = new_gs_list

    with pd.ExcelWriter("TEMP.xlsx", engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        df_struct.to_excel(writer, sheet_name="Struct_Para", index=False)

    print("🔍 'Parameter' cleaned and 'New_GS' extracted successfully in 'Struct_Para'.")

import pandas as pd
import re

def update_struct_gs_grouped():
    df = pd.read_excel("TEMP.xlsx", sheet_name="Struct_Para")

    # Ensure required columns are string
    df["New_GS"] = df["New_GS"].astype(str)
    df["Struct_Parameter"] = df["Struct_Parameter"].astype(str)
    df["GS Value"] = df["GS Value"].astype(str)
    df["MocPath"] = df["MocPath"].astype(str)
    df["Parameter"] = df["Parameter"].astype(str)

    # Create a grouping key
    df["group_key"] = df["MocPath"] + df["Parameter"]

    final_new_gs_dict = {}

    for group_key, group_df in df.groupby("group_key"):
        temp_new_gs = None

        for idx in group_df.index:
            struct_param = df.at[idx, "Struct_Parameter"].strip()
            gs_value = df.at[idx, "GS Value"].strip()
            current_new_gs = temp_new_gs if temp_new_gs is not None else df.at[idx, "New_GS"]

            # Match pattern: Struct_Parameter = <something>, <delimiter is , or ] or }>
            pattern = rf'({re.escape(struct_param)}\s*=\s*)([^,\]\}}]+)'

            def replacer(match):
                return match.group(1) + gs_value

            updated_new_gs = re.sub(pattern, replacer, current_new_gs, count=1)
            temp_new_gs = updated_new_gs  # carry forward

        # After all rows in group processed, apply last value to all
        final_new_gs_dict[group_key] = temp_new_gs

    # Assign final New_GS values to all rows in group
    for idx in df.index:
        df.at[idx, "New_GS"] = final_new_gs_dict[df.at[idx, "group_key"]]

    df.drop(columns=["group_key"], inplace=True)

    # Save changes back to TEMP.xlsx
    with pd.ExcelWriter("TEMP.xlsx", engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        df.to_excel(writer, sheet_name="Struct_Para", index=False)

    print("✅ 'New_GS' updated using row-wise Struct_Parameter replacement across group-wise structure.")


import pandas as pd

def handle_empty_containing_new_gs_struct_para():
    # Load Struct_Para sheet
    df = pd.read_excel("TEMP.xlsx", sheet_name="Struct_Para")

    # Track index groups with <empty> in New_GS and same Parameter
    grouped_data = []
    current_param = None
    temp_group = []

    for idx, row in df.iterrows():
        param = str(row["Parameter"]).strip()
        new_gs = str(row["New_GS"]).strip()

        if "<empty>" in new_gs:
            if current_param is None:
                current_param = param
            if param == current_param:
                temp_group.append(idx)
            else:
                # Flush previous group
                if temp_group:
                    grouped_data.append(temp_group)
                temp_group = [idx]
                current_param = param
        else:
            if temp_group:
                grouped_data.append(temp_group)
                temp_group = []
                current_param = None

    # Append remaining group
    if temp_group:
        grouped_data.append(temp_group)

    # Process each group
    for group in grouped_data:
        items = []
        for idx in group:
            struct_param = str(df.at[idx, "Struct_Parameter"]).strip()
            gs_val = str(df.at[idx, "GS Value"]).strip()
            items.append(f"{struct_param}={gs_val}")
        combined = "{" + ", ".join(items) + "}"
        for idx in group:
            df.at[idx, "New_GS"] = combined

    # Write back to file
    with pd.ExcelWriter("TEMP.xlsx", engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        df.to_excel(writer, sheet_name="Struct_Para", index=False)

    print("🧩 'New_GS' updated for all rows containing '<empty>' using Struct_Parameter=GS Value grouping.")

import pandas as pd

def update_struct_para_special_parameters():
    # Load Struct_Para sheet
    df = pd.read_excel("TEMP.xlsx", sheet_name="Struct_Para")

    # Define parameter-to-New_GS mapping
    param_to_newgs = {
        "daylightSavingTimeStartDate": '{dayRule="Sun>=8", month=MARCH, time="02:00"}',
        "daylightSavingTimeEndDate":   '{dayRule="Sun>=1", month=NOVEMBER, time="02:00"}',
        "rsrpPCellCandidatePcOffset":  '[{powerClass=PC_2, threshold2Offset=0}]',
        "rsrpPCellCandidateB2PcOffset": '[{powerClass=PC_2, threshold1Offset=0}]'
    }

    # Apply replacements
    for param, new_gs in param_to_newgs.items():
        mask = df["Parameter"] == param
        df.loc[mask, "New_GS"] = new_gs

    # Write updated Struct_Para back to TEMP.xlsx
    with pd.ExcelWriter("TEMP.xlsx", engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        df.to_excel(writer, sheet_name="Struct_Para", index=False)

    print("🎯 Special parameters updated in 'New_GS' column of Struct_Para.")


import pandas as pd


def update_plmnList_with_empty_gs_value():
    # Load Struct_Para tab
    df = pd.read_excel("TEMP.xlsx", sheet_name="Struct_Para")

    # Apply the condition: Parameter == "plmnList" AND GS Value is blank or empty
    mask = (df["Parameter"] == "plmnList") & (df["GS Value"].isna() | (df["GS Value"].astype(str).str.strip() == ""))

    # Set New_GS to []
    df.loc[mask, "New_GS"] = ""

    # Save back to TEMP.xlsx
    with pd.ExcelWriter("TEMP.xlsx", engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        df.to_excel(writer, sheet_name="Struct_Para", index=False)

    print("📬 'New_GS' set to '' for empty 'plmnList' entries.")

import pandas as pd
import re

def clean_new_gs_colon_prefix():
    # Load Struct_Para sheet
    df = pd.read_excel("TEMP.xlsx", sheet_name="Struct_Para")

    def trim_prefix(value):
        if pd.isna(value):
            return value
        value = str(value)
        if ":" in value:
            match = re.search(r'(\{|\[)', value)
            if match:
                return value[match.start():]  # Keep from { or [ onward
        return value

    # Apply the cleanup logic
    df["New_GS"] = df["New_GS"].apply(trim_prefix)

    # Write back to Excel
    with pd.ExcelWriter("TEMP.xlsx", engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        df.to_excel(writer, sheet_name="Struct_Para", index=False)

    print("🧼 Cleaned 'New_GS': removed colon-prefixed text before { or [.")


import pandas as pd


def wrap_new_gs_for_list_parameters():
    df = pd.read_excel("TEMP.xlsx", sheet_name="Struct_Para")

    # Filter where Parameter contains 'List' (case-insensitive)
    mask = df["Parameter"].astype(str).str.contains("List", case=False, na=False)

    for idx in df[mask].index:
        new_gs_val = str(df.at[idx, "New_GS"]).strip()

        if not new_gs_val or new_gs_val.lower() == "nan":
            df.at[idx, "New_GS"] = "[]"
        else:
            df.at[idx, "New_GS"] = f"[{new_gs_val}]"

    # Write back to TEMP.xlsx
    with pd.ExcelWriter("TEMP.xlsx", engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        df.to_excel(writer, sheet_name="Struct_Para", index=False)

    print("✅ 'New_GS' updated: empty values → '[]', others wrapped in [ ] where Parameter contains 'List'.")

import pandas as pd

def create_struct_para_uniq_updated():
    # Load Struct_Para sheet
    df = pd.read_excel("TEMP.xlsx", sheet_name="Struct_Para", dtype=str)

    # Select required columns
    uniq_df = df[["FULL_FDN", "Parameter", "New_GS"]].drop_duplicates()

    # Write to Struct_Para_Uniq tab
    with pd.ExcelWriter("TEMP.xlsx", engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        uniq_df.to_excel(writer, sheet_name="Struct_Para_Uniq", index=False)

    print("✅ 'Struct_Para_Uniq' tab created with unique (FULL_FDN, Parameter, New_GS) rows.")

import pandas as pd

def append_full_dump_to_script():
    # Load FULL_DUMP tab
    df = pd.read_excel("TEMP.xlsx", sheet_name="FULL_DUMP", dtype=str)

    # Drop rows where required columns are missing
    df = df.dropna(subset=["FDN_MO", "Parameter", "GS Value"])

    # Group by FDN_MO
    grouped = df.groupby("FDN_MO")

    script_lines = []

    for fdn, group in grouped:
        script_lines.append("set")
        script_lines.append(fdn.strip())
        for _, row in group.iterrows():
            param = str(row["Parameter"]).strip()
            value = str(row["GS Value"]).strip()
            script_lines.append(f"{param} : {value}")
        script_lines.append("")  # Optional: blank line between sets

    # Append to AREALITE_Delta_Script.txt
    with open("AREALITE_Delta_Script.txt", "a", encoding="utf-8") as f:
        f.write("\n".join(script_lines))
        f.write("\n")

    print("✅ Formatted FULL_DUMP data successfully appended to AREALITE_Delta_Script.txt.")

import pandas as pd

def append_struct_para_uniq_to_script():
    # Load Struct_Para_Uniq tab
    df = pd.read_excel("TEMP.xlsx", sheet_name="Struct_Para_Uniq", dtype=str)

    # Drop rows with missing values in critical columns
    df = df.dropna(subset=["FULL_FDN", "Parameter", "New_GS"])

    # Group by FULL_FDN
    grouped = df.groupby("FULL_FDN")

    script_lines = []

    for full_fdn, group in grouped:
        script_lines.append("set")
        script_lines.append(f'FDN : "{full_fdn.strip()}"')
        for _, row in group.iterrows():
            param = str(row["Parameter"]).strip()
            value = str(row["New_GS"]).strip()
            script_lines.append(f"{param} : {value}")
        script_lines.append("")  # Optional: spacing between blocks

    # Append to AREALITE_Delta_Script.txt
    with open("AREALITE_Delta_Script.txt", "a", encoding="utf-8") as f:
        f.write("\n".join(script_lines))
        f.write("\n")

    print("✅ Struct_Para_Uniq data successfully appended to AREALITE_Delta_Script.txt.")

# Main function to control the flow
def main():
    print("🔍 Step 1: Browsing and saving input files...")
    arealite_file, _ = browse_and_save_files()

    print("📄 Step 2: Converting AREALITE_Delta.txt to TEMP.xlsx...")
    df = create_temp_excel(arealite_file)

    print("🛠 Step 3: Processing CreateOnly and UniqueCreate tabs...")
    _, unique_df = process_create_section(df)

    print("📜 Step 4: Generating AREALITE_Delta_Script.txt...")
    generate_creation_script(unique_df)

    print("📑 Step 5: Creating Update tabs...")
    process_update_section(df)

    print("🗂 Step 6: Sorting 'CID_' sheets by 'MocPath'...")
    sort_cid_sheets()

    print("📥 Step 7: Extracting FDN & parameter blocks to FULL_DUMP tab...")
    parse_cli_dump_to_full_dump()

    print("✅ Automation complete! Output files:\n - TEMP.xlsx\n - AREALITE_Delta_Script.txt")

    print("🧹 Step 8: Splitting 'PARA_VALUE' into Parameter and Value...")
    split_full_dump_PARA_VALUE()

    print("🧩 Step 8: Adding FDN_MO to all CID_ sheets...")
    add_fdn_column_to_cid_tabs()

    print("🔗 Step 10: Adding 'FDN_MO_Parameter' to CID_ and FULL_DUMP tabs...")
    add_fdn_mo_parameter_column()

    print("🔄 Step 11: Mapping 'GS Value' from CID_ tabs into FULL_DUMP...")
    enrich_full_dump_with_gs_value()

    print("🧹 Step 12: Removing 'Not_FOUND' entries from FULL_DUMP...")
    remove_not_found_from_full_dump()

    print("🔠 Step 13: Converting GS Value integers to strings using FINAL tab...")
    convert_gs_value_to_string()

    print("🔳 Step 14: Wrapping GS Values in brackets for list-like parameters...")
    bracket_list_parameters()

    print("🧮 Step 15: Handling signed ranges/lists in GS Value column...")
    simplify_non_bracketed_ranges()

    print("🔄 Step 16: Conditional GS Value updates for '=' and ',' cases...")
    update_gs_value_conditionally()

    print("🔄 Step 16: Final GS Value cleanup using CLI_DUMP matching and quoting...")
    update_gs_value_conditionally()

    print("🔄 Step XX: Fixing GS Value based on CLI_DUMP with '=' and no quotes...")
    fix_gs_value_based_on_cli_dump()

    print("🧠 Step XX: Overriding GS Values for Special Parameters...")
    override_gs_value_from_special_category()

    print("🔍 Step XX: Filtering structured update parameters into 'Struct_Para' tab...")
    extract_struct_parameters_from_update()

    print("✅ 'FULL_FDN' created and Struct_Para sorted by 'FULL_FDN' + 'Parameter'.")
    update_full_fdn_and_sort()

    print("🧩 Step XX: Creating 'Struct_Parameter' from 'Parameter' column...")
    add_struct_parameter_column()

    print("🔄 Step XX: Resolving GS Value in Struct_Para from FINAL INT_String mapping...")
    resolve_struct_parameter_gs_value()

    print("🔍 Cleaning 'Parameter' and extracting New_GS from CLI_DUMP...")
    clean_parameter_and_extract_new_gs()

    print("✅ 'New_GS' updated using row-wise Struct_Parameter replacement across group-wise structure.")
    update_struct_gs_grouped()

    print("🧩 Aggregating Struct_Parameter=GS Value into New_GS where needed...")
    handle_empty_containing_new_gs_struct_para()

    print("🎯 Updating special Struct_Para parameters in New_GS...")
    update_struct_para_special_parameters()

    print("📬 Updating 'New_GS' for empty 'plmnList' values...")
    update_plmnList_with_empty_gs_value()

    print("🧼 Cleaned 'New_GS': removed colon-prefixed text before { or [.")
    clean_new_gs_colon_prefix()

    print("✅ 'New_GS' updated: empty values → '[]', others wrapped in [ ] where Parameter contains 'List'.")
    wrap_new_gs_for_list_parameters()

    print("✅ 'Struct_Para_Uniq' tab created with unique (FULL_FDN, Parameter, New_GS) rows.")
    create_struct_para_uniq_updated()

    print("✅ Formatted FULL_DUMP data successfully appended to AREALITE_Delta_Script.txt.")
    append_full_dump_to_script()

    print("✅ Struct_Para_Uniq data successfully appended to AREALITE_Delta_Script.txt.")
    append_struct_para_uniq_to_script()

# Run the script
if __name__ == "__main__":
    main()
