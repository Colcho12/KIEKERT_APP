import pandas as pd
from datetime import timedelta

"""Abre el archivo de Excel, y las hojas relacionadas con el Estado de Cuenta
RECIBE EL INPUT DEL GUI: ARCHIVO Y LOS DOS NOMBRES SELECCIONADOS"""
def StartEstadoCuenta(filepath, name_estado, name_aux):
    estado = pd.read_excel(filepath, sheet_name=name_estado)
    auxbanco = pd.read_excel(filepath, sheet_name=name_aux)

    auxbanco_cancelados = auxbanco[
        (auxbanco["Text"] == "cancelado") | (auxbanco["Text"] == "CANCELADO")
    ].copy()

    auxbanco_nocancelados = auxbanco[
        ~((auxbanco["Text"] == "cancelado") | (auxbanco["Text"] == "CANCELADO"))
    ].copy()

    return estado, auxbanco_cancelados, auxbanco_nocancelados

def Cleaning(estado,auxbanco):
    estado["FECHA"] = pd.to_datetime(estado["FECHA"], errors="coerce").dt.date
    auxbanco["Posting Date"] = pd.to_datetime(auxbanco["Posting Date"], errors="coerce").dt.date
    estado["Amount"] = estado["ABONOS"].fillna(0) - estado["CARGOS"].fillna(0)
    estado = estado[estado["FECHA"].notna()]
    return estado, auxbanco

def FirstSearch(estado, auxbanco):
    document_numbers = []  
    for index, row in estado.iterrows():
        
        match = auxbanco[
            (auxbanco["Posting Date"] == row["FECHA"]) &
            (auxbanco["Amount in doc. curr."] == row["Amount"]) &
            (~auxbanco["Used"])  
        ].head(1)  
        
        if not match.empty:
            document_number = match["Document Number"].iloc[0]
            auxbanco.loc[match.index, "Used"] = True  
        else:
            document_number = None  
        
        document_numbers.append(document_number)
    estadofinal = estado.copy()
    estadofinal["DOCUMENT NUMBER"] = document_numbers
    return estadofinal

def SecondSearch(estado, auxbanco):
    unmatched_rows = estado[estado["DOCUMENT NUMBER"].isna()]
    
    unique_amounts = auxbanco["Amount in doc. curr."].value_counts()
    
    unique_amounts = unique_amounts[unique_amounts == 1].index  # Only keep unique values
    for index, row in unmatched_rows.iterrows():
        # Search for a match where Amount matches Amount in doc. curr.
        match = auxbanco[
            (~auxbanco["Used"]) &
            (auxbanco["Amount in doc. curr."] == row["Amount"]) &
            (auxbanco["Amount in doc. curr."].isin(unique_amounts))  # Only consider unused matches
        ].head(1)

        if not match.empty:
            # Assign the Document Number and mark it as used
            document_number = match["Document Number"].iloc[0]
            auxbanco.loc[match.index, "Used"] = True
            estado.loc[index, "DOCUMENT NUMBER"] = document_number
    return estado

def ThirdSearch(estado, auxbanco):

    unmatched_rows = estado[estado["DOCUMENT NUMBER"].isna()]
    unmatched_rows = estado[estado["DOCUMENT NUMBER"].isna()]

    for index, row in unmatched_rows.iterrows():
        match = auxbanco[
            (~auxbanco["Used"]) &
            (auxbanco["Posting Date"] >= row["FECHA"] - timedelta(days=7)) &
            (auxbanco["Posting Date"] <= row["FECHA"] + timedelta(days=7)) &
            (auxbanco["Amount in doc. curr."] == row["Amount"]) &
            (~auxbanco["Used"])
        ].head(1)


        if not match.empty:
            document_number = match["Document Number"].iloc[0]
            auxbanco.loc[match.index, "Used"] = True
            estado.loc[index, "DOCUMENT NUMBER"] = document_number
        else:
            estado.loc[index, "DOCUMENT NUMBER"] = None

    return estado

def FourthSearch(estado, auxbanco):
    """
    Perform a search for unmatched rows where:
    - The absolute amount in "Estado de Cuenta" matches the auxiliary amount.
    - The sign is flipped (e.g., positive -> negative or vice versa).
    """
    # Identify rows where DOCUMENT NUMBER is still NA
    unmatched_rows = estado[estado["DOCUMENT NUMBER"].isna()]
    print(f"Number of unmatched rows before FourthSearch: {unmatched_rows.shape[0]}")  # Debugging

    for index, row in unmatched_rows.iterrows():
        match = auxbanco[
            (auxbanco["Amount in doc. curr."] == -row["Amount"]) &  # Flip the sign
            (~auxbanco["Used"])  # Ensure the row is not already used
        ].head(1)

        if not match.empty:
            document_number = match["Document Number"].iloc[0]
            auxbanco.loc[match.index, "Used"] = True
            estado.loc[index, "DOCUMENT NUMBER"] = document_number
        else:
            # Leave as NA if no match is found
            estado.loc[index, "DOCUMENT NUMBER"] = None

    print(f"Number of unmatched rows after FourthSearch: {estado['DOCUMENT NUMBER'].isna().sum()}")
    return estado

from datetime import timedelta

def find_consecutive_sum(values, target):
    """
    Find consecutive values (with at most one skip) that sum to the target.
    Returns the indices of the matching values if found, otherwise None.
    """
    n = len(values)
    for start in range(n):
        current_sum = 0
        skipped = False  # Allow one skip
        skip_index = -1  # Track the skipped index
        for end in range(start, n):
            current_sum += values[end]

            # Allow skipping one value if the sum overshoots
            if not skipped and abs(current_sum - target) > abs(target):
                skipped = True
                skip_index = end
                current_sum -= values[end]  # Remove the skipped value
                continue

            if current_sum == target:
                if skipped:
                    return list(range(start, skip_index)) + list(range(skip_index + 1, end + 1))
                return list(range(start, end + 1))
    return None

def find_all_consecutive_sums(values, target):
    """
    Return *all* subsets of consecutive values (with at most one skip)
    that sum exactly to `target`. Each subset is returned as a list of indices.
    """
    results = []
    n = len(values)
    for start in range(n):
        current_sum = 0
        skipped = False
        skip_index = -1
        for end in range(start, n):
            current_sum += values[end]
            if not skipped and abs(current_sum - target) > abs(target):
                # Attempt one skip
                skipped = True
                skip_index = end
                current_sum -= values[end]
                continue
            if current_sum == target:
                if skipped:
                    subset_indices = list(range(start, skip_index)) + list(range(skip_index+1, end+1))
                else:
                    subset_indices = list(range(start, end+1))
                results.append(subset_indices)
    return results

from collections import Counter, defaultdict

def remove_opposing_sign_pairs(values, indices):
    print(values)

    """
    Given parallel lists:
      values = [amt1, amt2, amt3, ...]
      indices = [idx1, idx2, idx3, ...]
    Remove any pairs x, -x (nonzero) from these lists, as many times as possible.
    Returns the 'cleaned' lists (new_values, new_indices) in the same *original order*.
    """
    c = Counter(values)
    
    # We'll track where each value's indices live
    pos_map = defaultdict(list)
    for v, i in zip(values, indices):
        pos_map[v].append(i)
    
    # Remove pairs
    for v in list(c.keys()):
        if v != 0 and -v in c:
            # Number of pairs we can remove is the min count of v and -v
            pairs_to_remove = min(c[v], c[-v])
            c[v]    -= pairs_to_remove
            c[-v]   -= pairs_to_remove
            # Remove from pos_map as well
            while pairs_to_remove > 0:
                pos_map[v].pop()
                pos_map[-v].pop()
                pairs_to_remove -= 1
            # If counts drop to zero, remove them from the dict
            if c[v] <= 0:
                del c[v]
            if c[-v] <= 0:
                del c[-v]

    # Rebuild the final lists but keep original order
    # We'll gather all remaining "value -> indices" then sort by index.
    remaining_pairs = []
    for v, cnt in c.items():
        for i in pos_map[v]:
            remaining_pairs.append((i, v))
    
    # Sort by original index i
    remaining_pairs.sort(key=lambda x: x[0])

    # Re-split into parallel lists
    cleaned_indices = [p[0] for p in remaining_pairs]
    cleaned_values  = [p[1] for p in remaining_pairs]
    return cleaned_values, cleaned_indices

def FifthSearch(estado, auxbanco, max_days=10):
    unmatched_aux_rows = auxbanco[~auxbanco["Used"]]
    print(f"Number of unmatched targets before FifthSearch: {unmatched_aux_rows.shape[0]}")

    for aux_index, aux_row in unmatched_aux_rows.iterrows():
        target_amount = aux_row["Amount in doc. curr."]
        target_date = aux_row["Posting Date"]
        document_number = aux_row["Document Number"]

        # Filter candidates in Estado de Cuenta within the date range
        candidates = estado[
            (estado["FECHA"] >= target_date - timedelta(days=max_days)) &
            (estado["FECHA"] <= target_date + timedelta(days=max_days)) &
            (estado["DOCUMENT NUMBER"].isna())  # Only unmatched
        ]
        if candidates.empty:
            continue

        candidate_values = candidates["Amount"].tolist()
        candidate_indices = candidates.index.tolist()

        # Find *all* possible subsets that sum to target_amount
        all_subsets = find_all_consecutive_sums(candidate_values, target_amount)
        if not all_subsets:
            continue  # No match at all for this Aux row

        # We'll try them in order:
        chosen_subset = None
        for subset_indices in all_subsets:
            # Build the amounts/indices for this subset
            matched_amounts  = [candidate_values[i] for i in subset_indices]
            matched_est_inds = [candidate_indices[i] for i in subset_indices]

            # Remove x,-x pairs
            cleaned_values, cleaned_indices = remove_opposing_sign_pairs(matched_amounts, matched_est_inds)
            
            # Check if the leftover amounts still sum to the same target
            if sum(cleaned_values) == target_amount:
                # Accept this subset
                chosen_subset = (cleaned_values, cleaned_indices)
                break  # We found a valid subset

        if chosen_subset is None:
            # Every subset we found had x, -x that caused the sum to change
            # => no valid match => skip
            continue

        # Otherwise, mark them as matched
        _, cleaned_indices = chosen_subset
        estado.loc[cleaned_indices, "DOCUMENT NUMBER"] = document_number
        auxbanco.at[aux_index, "Used"] = True

    print(f"Number of unmatched targets after FifthSearch: {auxbanco[~auxbanco['Used']].shape[0]}")
    return estado

def LastSearch(estado, auxbanco_cancelados):
    """
    Tries to fill any unmatched rows in 'estado' using 'auxbanco_cancelados'.
    Mark matched rows in auxbanco_cancelados as Used.
    """
    # First, clean the data similarly if needed
    _, auxbanco_cancelados = Cleaning(estado, auxbanco_cancelados)
    auxbanco_cancelados = auxbanco_cancelados.copy()
    
    # Ensure we have a 'Used' column
    if "Used" not in auxbanco_cancelados.columns:
        auxbanco_cancelados["Used"] = False

    unmatched_rows = estado[estado["DOCUMENT NUMBER"].isna()].copy()

    for index, row in unmatched_rows.iterrows():
        # For example, let's do a date +/- 7 days match + same amount
        match = auxbanco_cancelados[
            (~auxbanco_cancelados["Used"]) &
            (auxbanco_cancelados["Posting Date"] >= row["FECHA"] - timedelta(days=7)) &
            (auxbanco_cancelados["Posting Date"] <= row["FECHA"] + timedelta(days=7)) &
            (auxbanco_cancelados["Amount in doc. curr."] == row["Amount"]) &
            (~auxbanco_cancelados["Used"])
        ].head(1)

        if not match.empty:
            document_number = match["Document Number"].iloc[0]
            auxbanco_cancelados.loc[match.index, "Used"] = True
            estado.loc[index, "DOCUMENT NUMBER"] = document_number
        else:
            # If no match, keep it NaN
            estado.loc[index, "DOCUMENT NUMBER"] = None
    
    return estado


def MatchFechasMontos(estado_og, auxbanco_og):

    # LIMPIEZA DE DATOS
    estado,auxbanco = Cleaning(estado_og,auxbanco_og)

    auxbanco = auxbanco.copy()
    auxbanco["Used"] = False 

    estadouno = FirstSearch(estado, auxbanco)
    estadodos = SecondSearch(estadouno, auxbanco)
    estadotres = ThirdSearch(estadodos, auxbanco)
    estadocuatro = FourthSearch(estadotres, auxbanco)
    estadofinal = FifthSearch(estadocuatro, auxbanco)  # Sum search (final)

    return estadofinal

def run(filepath, name_estado, name_aux):
    estadocrudo, auxbanco_cancelados, auxbanco_nocancelados = StartEstadoCuenta(
        filepath, name_estado, name_aux
    )
    updated_estado = MatchFechasMontos(estadocrudo, auxbanco_nocancelados)
    updated_estado = LastSearch(updated_estado, auxbanco_cancelados)
    updated_estado.to_excel("./temp/fase1.xlsx", index=False, sheet_name="Updated Estado")
    print("File successfully saved as 'updated_estado_de_cuenta.xlsx'")