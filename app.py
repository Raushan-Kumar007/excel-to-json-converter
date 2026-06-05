import streamlit as st
import pandas as pd
import json

def build_tree(df):
    cols = [str(c) for c in df.columns]
    
    # Helper to find columns by keywords
    def find_col(keywords):
        for c in cols:
            if all(k.lower() in c.lower() for k in keywords):
                return c
        return None
        
    parent_id_col = find_col(['parent', 'id'])
    child_id_col = find_col(['child', 'id'])
    parent_name_col = find_col(['parent', 'name'])
    child_name_col = find_col(['child', 'name'])
    
    # If we have the 4 specific columns (Parent ID, Child ID, Parent Name, Child Name)
    if parent_id_col and child_id_col and parent_name_col and child_name_col:
        nodes = {}
        child_ids = set()
        
        df = df.where(pd.notnull(df), None)
        
        for _, row in df.iterrows():
            pid = row[parent_id_col]
            pname = row[parent_name_col]
            cid = row[child_id_col]
            cname = row[child_name_col]
            
            # Create parent node if it doesn't exist
            if pid is not None and pid not in nodes:
                nodes[pid] = {"Id": str(pid), "productName": str(pname) if pname else "", "children": []}
                
            # Create child node if it doesn't exist
            if cid is not None and cid not in nodes:
                nodes[cid] = {"Id": str(cid), "productName": str(cname) if cname else "", "children": []}
                
            # Link child to parent
            if pid is not None and cid is not None:
                child_node = nodes[cid]
                # Avoid duplicates
                if child_node not in nodes[pid]['children']:
                    nodes[pid]['children'].append(child_node)
                child_ids.add(cid)
                
        # Roots are nodes that are never listed as children
        roots = [n for pid, n in nodes.items() if pid not in child_ids]
        return roots

    else:
        # Fallback if columns don't match exactly, try to use Id, productName, ParentId
        df = df.where(pd.notnull(df), None)
        records = df.to_dict(orient="records")
        
        id_col = next((c for c in cols if str(c).lower() == 'id'), cols[0])
        name_col = next((c for c in cols if 'name' in str(c).lower()), cols[1] if len(cols) > 1 else cols[0])
        parent_col = next((c for c in cols if 'parent' in str(c).lower()), None)
        
        nodes = {}
        for r in records:
            node_id = r.get(id_col)
            if node_id is not None:
                nodes[node_id] = {
                    "Id": str(node_id),
                    "productName": str(r.get(name_col)) if r.get(name_col) else "",
                    "children": []
                }
                
        if not parent_col:
            return list(nodes.values())
            
        roots = []
        for r in records:
            node_id = r.get(id_col)
            pid = r.get(parent_col)
            
            if node_id is None:
                continue
                
            node = nodes[node_id]
            if pid is None or pid not in nodes or pid == node_id:
                roots.append(node)
            else:
                nodes[pid]['children'].append(node)
                
        return roots

st.set_page_config(page_title="Excel to JSON Converter", layout="wide")

st.markdown("<h1 style='text-align: center;'>Excel to JSON Converter</h1>", unsafe_allow_html=True)
st.markdown("<p style='text-align: center;'>Upload one or more Excel files to extract their data into a hierarchical JSON format.</p>", unsafe_allow_html=True)

uploaded_files = st.file_uploader("Choose Excel files", type=["xlsx", "xls"], accept_multiple_files=True)

if uploaded_files:
    try:
        # Read all Excel files and concatenate them
        dfs = []
        for file in uploaded_files:
            dfs.append(pd.read_excel(file))
            
        df = pd.concat(dfs, ignore_index=True)
        # Drop any completely blank rows and reset the index
        df.dropna(how='all', inplace=True)
        df.reset_index(drop=True, inplace=True)
        
        st.write("### Data Preview")
        st.dataframe(df, use_container_width=True)
        
        beautify = st.checkbox("Beautify JSON output", value=True)
        
        # Build the hierarchical tree automatically
        tree_data = build_tree(df)
        
        # If there's only one root, output it directly as an object
        if isinstance(tree_data, list) and len(tree_data) == 1:
            final_data = tree_data[0]
        else:
            final_data = tree_data
            
        if beautify:
            json_str = json.dumps(final_data, indent=4, default=str)
        else:
            json_str = json.dumps(final_data, default=str)
        
        st.write("### Extracted JSON")
        with st.container(height=500):
            st.code(json_str, language="json")
        # Generate an Excel file containing the hierarchical JSONs
        import io
        output_rows = []
        
        # Ensure roots is a list
        roots = tree_data if isinstance(tree_data, list) else [tree_data]
        
        for root in roots:
            output_rows.append({
                "Product Name": root.get("productName", ""),
                "Id": root.get("Id", ""),
                "Product Config Description": json.dumps(root, indent=4 if beautify else None, default=str)
            })
            
        out_df = pd.DataFrame(output_rows)
        
        with st.expander("Click to Preview Generated Excel"):
            st.dataframe(out_df, use_container_width=True)
        
        # Create an in-memory Excel file
        excel_buffer = io.BytesIO()
        with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
            out_df.to_excel(writer, index=False)
            
        # Provide download button
        file_name = f"{uploaded_files[0].name.split('.')[0]}_output.xlsx" if len(uploaded_files) > 1 else f"{uploaded_files[0].name.split('.')[0]}_output.xlsx"
        st.download_button(
            label="Download Excel",
            data=excel_buffer.getvalue(),
            file_name=file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
                            
    except Exception as e:
        st.error(f"Error reading the file: {e}")
