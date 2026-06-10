import streamlit as st
import pandas as pd
import json

def parse_level_hierarchy(df, cols, level_cols):
    roots = []
    last_nodes = {}
    
    for idx, row in df.iterrows():
        active_level = None
        node_name = None
        
        # Find which level this row corresponds to
        for i, c in enumerate(level_cols):
            val = row[c]
            if pd.notna(val) and str(val).strip() != "":
                active_level = i + 1
                node_name = str(val).strip()
                break
                
        if active_level is None:
            continue
            
        # Look for an explicit ID column
        id_col = next((c for c in cols if str(c).lower().strip() == 'id'), None)
        node_id = row[id_col] if id_col else None
        if pd.isna(node_id) or str(node_id).strip() == "":
            node_id = None
        else:
            node_id = str(node_id).strip()
                
        node = {
            "Id": node_id,
            "productName": node_name,
            "children": []
        }
        
        if active_level == 1:
            roots.append(node)
        else:
            parent_level = active_level - 1
            if parent_level in last_nodes:
                parent_node = last_nodes[parent_level]
                parent_node["children"].append(node)
            else:
                # If there's missing indentation, treat as root to avoid dropping
                roots.append(node)
                
        last_nodes[active_level] = node
        
    return roots

def build_tree(df):
    cols = [str(c) for c in df.columns]
    
    # Check if this is a "Level" based hierarchy file
    level_cols = [c for c in cols if 'level' in c.lower()]
    if len(level_cols) >= 2:
        return parse_level_hierarchy(df, cols, level_cols)
    
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
                parent_node = nodes[pid]
                
                # Avoid duplicates
                if child_node not in parent_node['children']:
                    parent_node['children'].append(child_node)
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
                parent_node = nodes[pid]
                parent_node['children'].append(node)
                
        return roots

st.set_page_config(page_title="Excel to JSON Converter", layout="wide")

st.markdown("<h1 style='text-align: center;'>Excel to JSON Converter</h1>", unsafe_allow_html=True)
st.markdown("<p style='text-align: center;'>Upload one or more Excel files to extract their data into a hierarchical JSON format.</p>", unsafe_allow_html=True)

uploaded_files = st.file_uploader("Choose Excel files", type=["xlsx", "xls"], accept_multiple_files=True)

if uploaded_files:
    try:
        all_roots = []
        all_dfs = []
        
        for file in uploaded_files:
            df = pd.read_excel(file)
            df.dropna(how='all', inplace=True)
            df.reset_index(drop=True, inplace=True)
            all_dfs.append(df)
            
            # Build the tree for this specific file
            tree_data = build_tree(df)
            
            # Combine roots
            if isinstance(tree_data, list):
                all_roots.extend(tree_data)
            else:
                all_roots.append(tree_data)
                
        # For the preview, we can concat all dfs (it might have lots of NaNs if formats differ, but it's just a preview)
        combined_df = pd.concat(all_dfs, ignore_index=True)
        
        st.write("### Combined Data Preview")
        st.dataframe(combined_df, use_container_width=True)
        
        beautify = st.checkbox("Beautify JSON output", value=True)
        
        # If there's only one root overall, we might format it as an object
        if len(all_roots) == 1:
            final_data = all_roots[0]
        else:
            final_data = all_roots
            
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
        
        def flatten_nodes_with_level(node_list, current_level=1):
            flat = []
            for n in node_list:
                flat.append((n, current_level))
                if n.get("children"):
                    flat.extend(flatten_nodes_with_level(n["children"], current_level + 1))
            return flat

        all_flat_nodes_with_level = flatten_nodes_with_level(all_roots)
        
        for node, level in all_flat_nodes_with_level:
            # Skip products that have no children
            if not node.get("children"):
                continue
                
            # Only include Level 2 products
            if level != 2:
                continue
                
            output_rows.append({
                "Product Name": node.get("productName", ""),
                "Id": node.get("Id", ""),
                "Product Config Description": json.dumps(node, indent=4 if beautify else None, default=str)
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
