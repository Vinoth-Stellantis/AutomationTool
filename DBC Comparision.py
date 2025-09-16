import sys
import cantools
import os
import xlsxwriter

def dbc_to_dict(db):
    """Convert a cantools database object into a dictionary for comparison."""
    dbc_dict = {}
    for msg in db.messages:
        msg_key = f"{msg.name}|{hex(msg.frame_id)}"
        msg_info = {
            "name": msg.name,
            "id": msg.frame_id,
            "senders": msg.senders,
            "receivers": msg.receivers,
            "signals": {sig.name: sig.name for sig in msg.signals}
        }
        dbc_dict[msg_key] = msg_info
    return dbc_dict

def compare_dbc(old_path, new_path, output_file):
    db_old = cantools.database.load_file(old_path)
    db_new = cantools.database.load_file(new_path)

    old_dict = dbc_to_dict(db_old)
    new_dict = dbc_to_dict(db_new)

    changes = []

    # --- Compare messages ---
    for msg_key, msg_old in old_dict.items():
        if msg_key not in new_dict:
            # Message Removed
            changes.append([
                msg_old["name"], hex(msg_old["id"]), "-", "-",
                ",".join(msg_old["senders"]), ",".join(msg_old["receivers"]),
                msg_old["name"], hex(msg_old["id"]), "-", "-",
                "-", "-",  # Tx/Rx only on old
                "Message Removed"
            ])
        else:
            msg_new = new_dict[msg_key]
            old_tx = ",".join(msg_old["senders"])
            old_rx = ",".join(msg_old["receivers"])
            new_tx = ",".join(msg_new["senders"])
            new_rx = ",".join(msg_new["receivers"])

            if old_tx != new_tx or old_rx != new_rx:
                changes.append([
                    msg_old["name"], hex(msg_old["id"]), "-", "-",
                    old_tx, old_rx,
                    msg_new["name"], hex(msg_new["id"]), "-", "-",
                    new_tx, new_rx,
                    "Tx/Rx Node Changed"
                ])
            # Signals removed
            for sig in msg_old["signals"]:
                if sig not in msg_new["signals"]:
                    changes.append([
                        msg_old["name"], hex(msg_old["id"]), sig, "-",
                        ",".join(msg_old["senders"]), ",".join(msg_old["receivers"]),
                        msg_old["name"], hex(msg_old["id"]), sig, "-",
                        "-", "-",  # Tx/Rx only on old
                        "Signal Removed"
                    ])
            # Signals added
            for sig in msg_new["signals"]:
                if sig not in msg_old["signals"]:
                    changes.append([
                        msg_new["name"], hex(msg_new["id"]), sig, "-",
                        "-", "-",  # Tx/Rx only on new
                        msg_new["name"], hex(msg_new["id"]), sig, "-",
                        ",".join(msg_new["senders"]), ",".join(msg_new["receivers"]),
                        "Signal Added"
                    ])

    # --- Check added messages ---
    for msg_key, msg_new in new_dict.items():
        if msg_key not in old_dict:
            changes.append([
                msg_new["name"], hex(msg_new["id"]), "-", "-",
                "-", "-",  # Tx/Rx only on new
                msg_new["name"], hex(msg_new["id"]), "-", "-",
                ",".join(msg_new["senders"]), ",".join(msg_new["receivers"]),
                "Message Added"
            ])

    columns = [
        "Old Msg Name", "Old Msg ID", "Old Signal", "Old Details",
        "Old Tx Node", "Old Rx Node",
        "New Msg Name", "New Msg ID", "New Signal", "New Details",
        "New Tx Node", "New Rx Node",
        "Comments"
    ]

    wb = xlsxwriter.Workbook(output_file)
    ws = wb.add_worksheet("DBC Comparison")

    bold = wb.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter'})
    red = wb.add_format({'font_color': 'red', 'align': 'center', 'valign': 'vcenter'})
    black = wb.add_format({'font_color': 'black', 'align': 'center', 'valign': 'vcenter'})
    border_fmt = wb.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter'})

    old_dbc = os.path.basename(old_path)
    new_dbc = os.path.basename(new_path)
    ws.merge_range(0, 0, 0, 5, f"Old DBC: {old_dbc}", bold)
    ws.merge_range(0, 6, 0, 11, f"New DBC: {new_dbc}", bold)
    ws.write(0, 12, "Comparison Results", bold)

    for col, col_name in enumerate(columns):
        ws.write(1, col, col_name, bold)
        ws.set_column(col, col, 22)

    row_idx = 2

    def write_cell(r, c, value, fmt=border_fmt):
        ws.write(r, c, value if value not in [None, ""] else "-", fmt)

    def rich_write(r, c, values, added=None, removed=None):
        if isinstance(values, str):
            values = [v.strip() for v in values.split(',') if v.strip()]
        if not values:
            ws.write(r, c, "-", border_fmt)
            return
        if len(values) == 1:
            n = values[0]
            fmt = black
            if (added and n in added) or (removed and n in removed):
                fmt = red
            ws.write(r, c, n, fmt)
            return
        parts = []
        for i, n in enumerate(values):
            fmt = black
            if added and n in added:
                fmt = red
            if removed and n in removed:
                fmt = red
            parts.extend([fmt, n])
            if i < len(values) - 1:
                parts.extend([black, ", "])
        ws.write_rich_string(r, c, *parts, border_fmt)

    for row_data in changes:
        comment = row_data[12]

        # Default write
        for c, val in enumerate(row_data):
            write_cell(row_idx, c, val)

        # Apply coloring rules
        if comment == "Message Removed":
            for c in range(0, 6):  # Old side
                ws.write(row_idx, c, row_data[c], black)
            for c in range(6, 12):  # New side
                ws.write(row_idx, c, row_data[c], red)
        elif comment == "Message Added":
            for c in range(0, 6):  # Old side
                ws.write(row_idx, c, row_data[c], red)
            for c in range(6, 12):  # New side
                ws.write(row_idx, c, row_data[c], black)
        elif comment == "Signal Removed":
            # Only signal in new side is red; Msg name & ID remain black
            ws.write(row_idx, 0, row_data[0], black)
            ws.write(row_idx, 1, row_data[1], black)
            ws.write(row_idx, 2, row_data[2], black)
            ws.write(row_idx, 6, row_data[6], black)
            ws.write(row_idx, 7, row_data[7], black)
            ws.write(row_idx, 8, row_data[8], red)  # signal removed
        elif comment == "Signal Added":
            # Only signal in old side is red; Msg name & ID remain black
            ws.write(row_idx, 0, row_data[0], black)
            ws.write(row_idx, 1, row_data[1], black)
            ws.write(row_idx, 2, row_data[2], red)  # signal added
            ws.write(row_idx, 6, row_data[6], black)
            ws.write(row_idx, 7, row_data[7], black)
            ws.write(row_idx, 8, row_data[8], black)
        elif comment == "Tx/Rx Node Changed":
            old_tx_list = row_data[4].split(",") if row_data[4] != "-" else []
            old_rx_list = row_data[5].split(",") if row_data[5] != "-" else []
            new_tx_list = row_data[10].split(",") if row_data[10] != "-" else []
            new_rx_list = row_data[11].split(",") if row_data[11] != "-" else []

            added_tx = [n for n in new_tx_list if n not in old_tx_list]
            removed_tx = [n for n in old_tx_list if n not in new_tx_list]
            added_rx = [n for n in new_rx_list if n not in old_rx_list]
            removed_rx = [n for n in old_rx_list if n not in new_rx_list]

            rich_write(row_idx, 4, old_tx_list, removed=removed_tx)
            rich_write(row_idx, 5, old_rx_list, removed=removed_rx)
            rich_write(row_idx, 10, new_tx_list, added=added_tx)
            rich_write(row_idx, 11, new_rx_list, added=added_rx)

        row_idx += 1

    ws.autofilter(1, 0, row_idx - 1, len(columns) - 1)
    ws.freeze_panes(2, 0)
        # === Add borders before closing ===
    thin_border = wb.add_format({'border': 1})
    thick_border = wb.add_format({'border': 2})

    # Apply thin border to all cells inside the table
    ws.conditional_format(0, 0, row_idx - 1, len(columns) - 1,
                          {'type': 'no_errors', 'format': thin_border})

    # Apply thick outside border to the entire comparison range
    ws.conditional_format(0, 0, row_idx - 1, len(columns) - 1,
                          {'type': 'no_errors', 'format': thick_border})

    wb.close()
    print(f"Comparison complete. Differences saved in: {output_file}")

if __name__ == "__main__":
    if len(sys.argv) != 4:
        print("Usage: python dbc_compare_excel.py <old_dbc> <new_dbc> <output_excel>")
        sys.exit(1)

    old_dbc = sys.argv[1]
    new_dbc = sys.argv[2]
    out_excel = sys.argv[3]

    compare_dbc(old_dbc, new_dbc, out_excel)
