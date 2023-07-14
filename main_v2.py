# -*- coding: utf-8 -*-
import random
import streamlit as st
import pandas as pd
from ipfn import ipfn


st.title("Matrix Balancing Tool")
file_container = st.container()
sheets_container = st.container()
table_container = st.container()
check_field_container = st.container()
result_container = st.container()
st.session_state.create_new_seed = False
st.session_state.compare = None
st.session_state.df_compare = None
st.session_state.button_read_table = False
st.session_state.df_seed = None
st.session_state.df_A = None
st.session_state.df_B = None


def get_sheets_and_rows(file):
    excel_file = pd.ExcelFile(uploaded_file)
    excel_sheets = excel_file.sheet_names
    with sheets_container:
        col1, col2 = st.columns(2)

        with col1:
            sheet_A = st.selectbox("Choose the Target A sheet", excel_sheets)
            sheet_B = st.selectbox("Choose the Target B sheet", excel_sheets)
            sheet_S = st.selectbox("Choose the SEED sheet", excel_sheets)

        with col2:
            row_A = st.number_input("Target A table start row", min_value=1, step=1, value=1, key="A")
            row_B = st.number_input("Target B table start row", min_value=1, step=1, value=1, key="B")
            row_S = st.number_input("Seed table start row", min_value=1, step=1, value=1, key="S")

    return sheet_A, sheet_B, sheet_S, int(row_A), int(row_B), int(row_S)


# @st.cache_data
def get_number_of_field(df):
    i = 0
    for column in df.columns:

        if column == "Year":
            i += 1
        else:

            try:
                pd.to_numeric(df[column])
                if "Year" in df.columns and i != 0:
                    break
                else:
                    int(column)
                    break

            except ValueError:
                i += 1

    return i


# @st.cache_data
def rename_industry_column_name(x):
    if type(x) == str:
        x = x.strip()

    if x == "Industry":
        x = "Industry sector"

    return x


def validate_field_name(target_field, seed_field):
    with table_container:
        for field in target_field:
            if field not in seed_field:
                st.write(f"{field} not found in SEED table. Please check and try again.")
                return False

    return True


# @st.cache_data
def drop_unmatch_rows(df_seed, df_A, df_B):
    df_seed = pd.DataFrame(df_seed.stack())
    df_seed = df_seed.reset_index()

    for column in df_seed.columns[:-1]:
        if column in df_A.index.names:
            df_seed = df_seed[df_seed[column].isin(df_A.index.get_level_values(column).tolist())]
        elif column in df_B.index.names:
            df_seed = df_seed[df_seed[column].isin(df_B.index.get_level_values(column).tolist())]
        elif column == "Year":
            df_seed = df_seed[df_seed[column].isin(df_A.columns.tolist())]

    for index in df_A.index.names:
        df_A = df_A.query(f'`{index}` in {df_seed[index].tolist()}')

    for index in df_B.index.names:
        df_B = df_B.query(f'`{index}` in {df_seed[index].tolist()}')

    df_A = df_A.stack()
    df_B = df_B.stack()

    return df_seed, df_A, df_B


# @st.cache_data
def validate_field_item(df_seed, df_A, df_B):
    df_seed = pd.DataFrame(df_seed.stack())
    df_seed = df_seed.reset_index()

    df_compare = pd.DataFrame(columns=["seed", "target_A", "target_B"])
    compare = True

    for column in df_seed.columns[:-1]:
        add_data = pd.DataFrame()
        seed_field_item = sorted(df_seed[column].unique().tolist())
        add_data = pd.concat([add_data,
                              pd.DataFrame(data={"seed": seed_field_item},
                                           index=[column] * len(seed_field_item))], axis=1)
        if column in df_A.index.names:
            df_A_field_item = sorted(df_A.index.get_level_values(column).dropna().unique().tolist())
            if df_A_field_item != seed_field_item:
                compare = False
            while len(df_A_field_item) < len(seed_field_item):
                df_A_field_item.append("")
            while len(seed_field_item) < len(df_A_field_item):
                seed_field_item.append("")
            add_data = pd.concat([add_data,
                                  pd.DataFrame(data={"target_A": df_A_field_item},
                                               index=[column] * len(df_A_field_item))], axis=1)

        if column in df_B.index.names:
            df_B_field_item = sorted(df_B.index.get_level_values(column).dropna().unique().tolist())
            if df_B_field_item != seed_field_item:
                compare = False
            while len(df_B_field_item) < len(seed_field_item):
                df_B_field_item.append("")
            while len(seed_field_item) < len(df_B_field_item):
                seed_field_item.append("")
            add_data = pd.concat([add_data,
                                  pd.DataFrame(data={"target_B": df_B_field_item},
                                               index=[column] * len(df_B_field_item))], axis=1)

        df_compare = pd.concat([df_compare, add_data])

    return compare, df_compare


# @st.cache_data
def create_new_seed_table(df_A, df_B, field_name_A, field_name_B):

    index_names = []
    list_field_item = []

    for field in field_name_A:
        list_field_item.append(df_A[field].dropna().unique().tolist())
        index_names.append(field)

    for field in field_name_B:
        if field not in field_name_A:
            list_field_item.append(df_B[field].dropna().unique().tolist())
            index_names.append(field)

    df_A = df_A.set_index(field_name_A)
    year_columns = df_A.columns.tolist()

    index = pd.MultiIndex.from_product(list_field_item).set_names(index_names)

    df_seed = pd.DataFrame(index=index, columns=year_columns).fillna(1).reset_index()

    return df_seed


def read_table(file, sheet_A, sheet_B, sheet_S, row_A, row_B, row_S):
    # read tables by sheetname, header set to None as the start row is different
    df_A = pd.read_excel(file, sheet_name=sheet_A, skiprows=row_A - 1).dropna(axis=1, how="all").dropna(axis=0,
                                                                                                        how="all")
    df_B = pd.read_excel(file, sheet_name=sheet_B, skiprows=row_B - 1).dropna(axis=1, how="all").dropna(axis=0,
                                                                                                        how="all")

    # get number of field
    field_A = get_number_of_field(df_A)
    field_B = get_number_of_field(df_B)

    # replace "Industry" to "Industry sector" in columns
    df_A.columns = list(map(lambda name: rename_industry_column_name(name), df_A.columns))
    df_B.columns = list(map(lambda name: rename_industry_column_name(name), df_B.columns))

    # collect all the field name
    field_name_A = df_A.columns[:field_A].tolist()
    field_name_B = df_B.columns[:field_B].tolist()

    if not st.session_state.create_new_seed:

        df_seed = pd.read_excel(file, sheet_name=sheet_S, skiprows=row_S - 1).dropna(axis=1, how="all").\
            dropna(axis=0, how="all")

    else:
        df_seed = create_new_seed_table(df_A, df_B, field_name_A, field_name_B)
        writer = pd.ExcelWriter(uploaded_file, engine='openpyxl', mode='a',
                                if_sheet_exists='new')
        df_seed.to_excel(writer, sheet_name='NEW_SEED', engine='openpyxl', index=False)
        writer.close()
        # st.write(df_seed.astype(str))

    # get number of field for seed table
    field_S = get_number_of_field(df_seed)
    # replace "Industry" to "Industry sector" in columns for seed table
    df_seed.columns = list(map(lambda name: rename_industry_column_name(name), df_seed.columns))
    # collect all the field name for seed table
    field_name_S = df_seed.columns[:field_S].tolist()
    field_name = list(set(field_name_S + field_name_A + field_name_B + ["Year"]))

    # set index for each tables
    df_seed = df_seed.set_index(field_name_S)
    df_A = df_A.set_index(field_name_A)
    df_B = df_B.set_index(field_name_B)

    # by default the columns for target table is Year, transform the year data type to integer
    df_A.columns.name = "Year"
    df_B.columns.name = "Year"
    df_A.columns = pd.to_numeric(df_A.columns)
    df_B.columns = pd.to_numeric(df_B.columns)

    for field in field_name:
        if field not in field_name_S:
            df_seed.columns.name = field
            field_name_S.append(field)
            break

    # check if seed columns is year
    if df_seed.columns.name == "Year":
        df_seed.columns = pd.to_numeric(df_seed.columns)

    # validate field name
    with table_container:
        if validate_field_name(field_name_A, field_name_S) and \
                validate_field_name(field_name_B, field_name_S) and \
                df_A.columns.tolist() == df_B.columns.tolist():

            st.write(f"Target A: {' x '.join(field_name_A)} x Year")
            st.write(f"Target B: {' x '.join(field_name_B)} x Year")
            st.write(f"Seed: {' x '.join(field_name_S)}")

            if st.session_state.compare is None:
                # check field item of each table
                compare, df_compare = validate_field_item(df_seed, df_A, df_B)
                st.session_state["compare"] = compare
                st.session_state["df_compare"] = df_compare

            # if all the field item is matched then proceed
            if st.session_state["compare"]:
                with check_field_container:
                    st.info("Field items are matched.")
                writer = pd.ExcelWriter(uploaded_file, engine='openpyxl', mode='a',
                                        if_sheet_exists='new')
                st.session_state["df_compare"].to_excel(writer, sheet_name='CHECK_TABLES_OK', engine='openpyxl')
                writer.close()
            else:
                with check_field_container:
                    st.warning("Field item does not match. Please check the file.")
                writer = pd.ExcelWriter(uploaded_file, engine='openpyxl', mode='a',
                                        if_sheet_exists='new')
                st.session_state["df_compare"].to_excel(writer, sheet_name='CHECK_TABLES_NOT_OK', engine='openpyxl')
                writer.close()

            # drop unmatch rows
            df_seed, df_A, df_B = drop_unmatch_rows(df_seed, df_A, df_B)

            with st.expander("Click here to see table"):
                # st.write("Compare", df_compare.astype(str))
                # col1, col2 = st.columns(2)
                #
                # with col1:
                st.write("Target A", df_A)
                # with col2:
                st.write("Target B", df_B)

                st.write("SEED Table", df_seed)

            st.write("New tables loaded.")

            if st.download_button("Download Check Results", uploaded_file, file_name=uploaded_file.name):
                pass
            if st.session_state.create_new_seed:
                if st.download_button("Download New Seed Table", uploaded_file, file_name=uploaded_file.name):
                    pass

            return df_seed, df_A, df_B

        elif df_A.columns.tolist() != df_B.columns.tolist():
            st.write("Year columns in Target A does not match with Years column in Target B")
            return None, None, None

        else:
            return None, None, None


# @st.cache_data
def format_result_table(df_result, df_seed_index):
    # format result table if N=3
    if len(df_seed_index) == 3:
        df_result = df_result.set_index(df_seed_index).unstack(level=2).droplevel(0, axis=1)

        df_grand_total_row = df_result.sum()

        df_subtotal = df_result.groupby([df_seed_index[0]], level=0).sum()

        # add subtotal row
        for x in df_subtotal.index.tolist():
            df_result.loc[(x, "Total"), :] = df_subtotal.loc[x, :]
            df_result = df_result.sort_index(level=0)

        # add grand total row
        if df_result.columns.name != "Year":
            df_result.insert(0, "Total industry", df_result.sum(axis=1))
        else:
            df_result.loc[("Grand Total", "Grand Total"), :] = df_grand_total_row

    else:
        df_result = df_result.set_index(df_seed_index).unstack(level=-1).droplevel(0, axis=1)

    return df_result


def generate_results(df_seed, df_A, df_B):
    # save the initial seed index for later formatting
    df_seed_index = df_seed.columns.tolist()[0:-1]

    aggregates = [df_A, df_B]
    dimensions = [list(df_A.index.names), list(df_B.index.names)]

    with result_container:
        col1, col2, col3 = st.columns(3)

        with col1:
            convergence_rate = st.number_input("Convergence rate", value=1e-5, step=1e-5, format="%.f", key="conv")
        with col2:
            rate_tolerance = st.number_input("Tolerance rate", value=1e-8, step=1e-8, format="%.f", key="tol")
        with col3:
            max_iteration = st.number_input("Maximum iteration", step=1, value=500, key="iter")

        if st.button("Generate Results", random.randint(0, 100000)):
            try:
                IPF = ipfn(df_seed, aggregates, dimensions, weight_col=0, verbose=2,
                           convergence_rate=convergence_rate, rate_tolerance=rate_tolerance,
                           max_iteration=max_iteration)
                df, flag, df_iteration = IPF.iteration()
                st.write(df_seed)

                iteration = max(df_iteration.index)
                conv_rate = df_iteration.iat[iteration, 0]

                st.write(f"Number of Iteration: {iteration + 1}")
                st.write(f"Convergence rate: {conv_rate}")

                with st.spinner("Saving results..."):

                    writer = pd.ExcelWriter(uploaded_file, engine='openpyxl', mode='a',
                                            if_sheet_exists='new')
                    df = df.rename(columns={0: "Value"})
                    df.to_excel(writer, sheet_name='Results', index=False, engine='openpyxl')

                    df_result = format_result_table(df, df_seed_index)
                    df_result.to_excel(writer, sheet_name="Results_formatted", merge_cells=False, engine='openpyxl')

                    writer.close()

                st.success("Results saved.")

                with result_container:
                    if st.download_button("Download Results", uploaded_file, file_name=uploaded_file.name):
                        pass

            except Exception as e:
                st.exception(e)


with file_container:
    st.header("Choose an input file")

    uploaded_file = st.file_uploader("Choose an excel file", type="xlsx")

if uploaded_file is None:
    st.stop()

sheet_A, sheet_B, sheet_S, row_A, row_B, row_S = get_sheets_and_rows(uploaded_file)

with sheets_container:
    create_new_seed = st.checkbox("Create new seed table")
    button_read = st.button("Read tables")

if button_read:
    st.session_state.button_read_table = True
    if create_new_seed:
        st.session_state.create_new_seed = True
    else:
        st.session_state.create_new_seed = False

if st.session_state.button_read_table:
    with table_container:
        df_seed, df_A, df_B = read_table(uploaded_file, sheet_A, sheet_B, sheet_S, row_A, row_B, row_S)

        st.session_state["df_seed"] = df_seed
        st.session_state["df_A"] = df_A
        st.session_state["df_B"] = df_B

if st.session_state["df_seed"] is not None and \
        st.session_state["df_A"] is not None and st.session_state["df_B"] is not None:
    generate_results(st.session_state["df_seed"], st.session_state["df_A"], st.session_state["df_B"])

st.write(uploaded_file.name)
