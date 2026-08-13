import streamlit as st

from serials_core import (
    HardwareItem,
    SoftwareItem,
    all_hardware_serials,
    all_software_keys,
    clean_hardware_items,
    clean_software_items,
    create_serials_excel,
    find_duplicates,
    hardware_serial_count,
    parse_bulk_hardware,
    parse_pasted_values,
    safe_filename_part,
    software_key_count,
)


def _init_state():
    st.session_state.setdefault("hardware_count", 1)
    st.session_state.setdefault("software_count", 1)
    st.session_state.setdefault("include_hardware", True)
    st.session_state.setdefault("include_software", False)


def _reset_report():
    keep = {"include_hardware": True, "include_software": False, "hardware_count": 1, "software_count": 1}
    for key in list(st.session_state.keys()):
        if key.startswith(("client_", "order_", "po_", "hw_", "sw_", "bulk_", "generated_")):
            st.session_state.pop(key, None)
    for key, value in keep.items():
        st.session_state[key] = value


def _render_hardware() -> list[HardwareItem]:
    st.subheader("Hardware serial numbers")
    entry_method = st.radio(
        "How do you want to enter hardware?",
        ["Easy entry", "Bulk paste from Excel"],
        horizontal=True,
        key="hw_entry_method",
        label_visibility="collapsed",
    )

    if entry_method == "Bulk paste from Excel":
        st.caption("Copy rows from Excel in this order: Product | Model | Serial Number. Two-column Product | Serial pastes also work.")
        raw = st.text_area(
            "Paste hardware rows",
            key="bulk_hardware",
            height=220,
            placeholder="Dell Latitude 7450\tLatitude 7450\tABC1234\nDell Latitude 7450\tLatitude 7450\tABC1235",
        )
        items, ignored = parse_bulk_hardware(raw)
        count = hardware_serial_count(items)
        if raw.strip():
            if count:
                st.success(f"{count} serial number{'s' if count != 1 else ''} detected across {len(items)} product{'s' if len(items) != 1 else ''}.")
            if ignored:
                st.warning(f"{len(ignored)} row{'s were' if len(ignored) != 1 else ' was'} skipped. Make sure each row has at least Product and Serial Number columns.")
        return items

    items: list[HardwareItem] = []
    for i in range(st.session_state.hardware_count):
        with st.container(border=True):
            st.markdown(f"**Hardware item {i + 1}**")
            c1, c2 = st.columns(2)
            with c1:
                product = st.text_input("Product", key=f"hw_product_{i}", placeholder="Dell Latitude 7450")
            with c2:
                model = st.text_input("Model number", key=f"hw_model_{i}", placeholder="Latitude 7450")
            serial_text = st.text_area(
                "Serial numbers",
                key=f"hw_serials_{i}",
                height=120,
                placeholder="Paste one per line, from Excel, or comma-separated...",
            )
            serials = parse_pasted_values(serial_text)
            duplicates = find_duplicates(serials)
            if serials and not duplicates:
                st.caption(f"✓ {len(serials)} serial number{'s' if len(serials) != 1 else ''} detected")
            if duplicates:
                st.warning("Duplicate serial number" + ("s" if len(duplicates) != 1 else "") + ": " + ", ".join(duplicates[:8]))
            items.append(HardwareItem(product=product, model=model, serials=serials))

    add_col, remove_col, _ = st.columns([1.4, 1.4, 4])
    with add_col:
        if st.button("+ Add another product", use_container_width=True):
            st.session_state.hardware_count += 1
            st.rerun()
    with remove_col:
        if st.session_state.hardware_count > 1 and st.button("Remove last", use_container_width=True):
            last = st.session_state.hardware_count - 1
            for suffix in ("product", "model", "serials"):
                st.session_state.pop(f"hw_{suffix}_{last}", None)
            st.session_state.hardware_count -= 1
            st.rerun()
    return items


def _render_software() -> list[SoftwareItem]:
    st.subheader("Software license keys")
    items: list[SoftwareItem] = []
    for i in range(st.session_state.software_count):
        with st.container(border=True):
            st.markdown(f"**Software item {i + 1}**")
            c1, c2 = st.columns(2)
            with c1:
                software = st.text_input("Software", key=f"sw_name_{i}", placeholder="Microsoft Office")
            with c2:
                license_type = st.text_input("License type", key=f"sw_type_{i}", placeholder="Perpetual / Subscription / MAK")
            key_text = st.text_area(
                "License keys",
                key=f"sw_keys_{i}",
                height=120,
                placeholder="Paste one per line, from Excel, or comma-separated...",
            )
            keys = parse_pasted_values(key_text)
            duplicates = find_duplicates(keys)
            if keys and not duplicates:
                st.caption(f"✓ {len(keys)} license key{'s' if len(keys) != 1 else ''} detected")
            if duplicates:
                st.warning("Duplicate license key" + ("s" if len(duplicates) != 1 else "") + ": " + ", ".join(duplicates[:5]))
            items.append(SoftwareItem(software=software, license_type=license_type, license_keys=keys))

    add_col, remove_col, _ = st.columns([1.4, 1.4, 4])
    with add_col:
        if st.button("+ Add another software item", use_container_width=True):
            st.session_state.software_count += 1
            st.rerun()
    with remove_col:
        if st.session_state.software_count > 1 and st.button("Remove last software", use_container_width=True):
            last = st.session_state.software_count - 1
            for prefix in ("sw_name_", "sw_type_", "sw_keys_"):
                st.session_state.pop(f"{prefix}{last}", None)
            st.session_state.software_count -= 1
            st.rerun()
    return items


def main():
    st.set_page_config(page_title="Safari Micro | Serials & Licenses", page_icon="📄", layout="wide")
    _init_state()

    st.image("https://safarimicro.com/wp-content/uploads/2022/01/SafariMicro-Color-with-Solid-Icon-Copy.png", width=230)
    st.title("Serial Numbers & License Report")
    st.caption("Create a clean, customer-ready Excel report without fighting with formatting.")

    st.subheader("1. Customer information")
    c1, c2, c3 = st.columns(3)
    with c1:
        client_name = st.text_input("Client name", key="client_name")
    with c2:
        order_number = st.text_input("Safari order #", key="order_number")
    with c3:
        client_po = st.text_input("Customer PO", key="po_number")

    st.subheader("2. What are you sending?")
    option1, option2, _ = st.columns([1.5, 1.5, 4])
    with option1:
        include_hardware = st.checkbox("Hardware serial numbers", key="include_hardware")
    with option2:
        include_software = st.checkbox("Software license keys", key="include_software")

    hardware_items: list[HardwareItem] = []
    software_items: list[SoftwareItem] = []

    if include_hardware:
        st.divider()
        hardware_items = _render_hardware()
    if include_software:
        st.divider()
        software_items = _render_software()

    hardware_items = clean_hardware_items(hardware_items)
    software_items = clean_software_items(software_items)

    st.divider()
    st.subheader("3. Create the report")

    hw_count = hardware_serial_count(hardware_items)
    sw_count = software_key_count(software_items)
    hw_duplicates = find_duplicates(all_hardware_serials(hardware_items))
    sw_duplicates = find_duplicates(all_software_keys(software_items))

    problems: list[str] = []
    if not client_name.strip():
        problems.append("Add the client name.")
    if not order_number.strip():
        problems.append("Add the Safari order number.")
    if not include_hardware and not include_software:
        problems.append("Select hardware, software, or both.")
    if include_hardware and hw_count == 0:
        problems.append("Add at least one hardware serial number.")
    if include_software and sw_count == 0:
        problems.append("Add at least one software license key.")
    if hw_duplicates:
        problems.append("Remove duplicate hardware serial numbers: " + ", ".join(hw_duplicates[:8]))

    if problems:
        st.warning("Please finish the items below before creating the report:\n\n" + "\n".join(f"• {problem}" for problem in problems))
    elif sw_duplicates:
        st.info("The report can be created, but duplicate software keys were detected. This can be valid for some licensing programs, so please confirm before sending.")
    else:
        summary = []
        if hw_count:
            summary.append(f"{hw_count} hardware serial number{'s' if hw_count != 1 else ''}")
        if sw_count:
            summary.append(f"{sw_count} software license key{'s' if sw_count != 1 else ''}")
        st.success("Report looks good — " + " and ".join(summary) + ".")

    left, middle, _ = st.columns([1.8, 1, 5])
    with left:
        generate = st.button("Create Excel report", type="primary", use_container_width=True, disabled=bool(problems))
    with middle:
        st.button("Start over", on_click=_reset_report, use_container_width=True)

    if generate:
        excel_data = create_serials_excel(client_name, order_number, client_po, hardware_items, software_items)
        client_part = safe_filename_part(client_name, "Client")
        order_part = safe_filename_part(order_number, "Order")
        st.session_state.generated_excel = excel_data
        st.session_state.generated_filename = f"SafariMicro_{client_part}_{order_part}_Serials.xlsx"

    if st.session_state.get("generated_excel"):
        st.download_button(
            "Download customer report",
            data=st.session_state.generated_excel,
            file_name=st.session_state.generated_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )
        st.caption("Nothing is saved by this app after you leave the page; the Excel file is generated in memory for download.")


if __name__ == "__main__":
    main()
