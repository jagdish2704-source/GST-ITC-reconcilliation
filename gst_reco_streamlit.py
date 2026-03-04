import os
import tempfile
import streamlit as st
import pandas as pd

# import core functions from existing module
from gst_reco_app import (
    init_db,
    verify_login,
    create_user,
    delete_user,
    get_all_users,
    reconcile,
    generate_correction_report,
    create_default_change_heading_file,
)


# initialize database and config when the app starts
init_db()
create_default_change_heading_file()


# helper callbacks for progress and logging

def make_logger():
    log_lines = []

    def _log(msg: str):
        log_lines.append(msg)
        # update placeholder text
        log_placeholder.text("\n".join(log_lines))

    log_placeholder = st.empty()
    return _log


def main():
    if "logged_in" not in st.session_state:
        st.session_state.logged_in = False
        st.session_state.username = None
        st.session_state.role = None

    st.title("GST ITC Reconciliation - Streamlit Edition")

    if not st.session_state.logged_in:
        login_ui()
    else:
        dashboard_ui()


def login_ui():
    st.subheader("Please sign in")
    user = st.text_input("User ID")
    pwd = st.text_input("Password", type="password")
    if st.button("Login"):
        role = verify_login(user, pwd)
        if role:
            st.session_state.logged_in = True
            st.session_state.username = user
            st.session_state.role = role
            st.experimental_rerun()
        else:
            st.error("Invalid user ID or password.")


def dashboard_ui():
    # sidebar info
    st.sidebar.write(f"Logged in as: **{st.session_state.username}** ({st.session_state.role})")
    if st.session_state.role == "ADMIN":
        admin_panel()

    st.header("Reconciliation Inputs")
    taxpayer = st.text_input("Taxpayer GSTIN", key="taxpayer")
    period = st.text_input("Reconciliation Period", key="period")
    tolerance = st.selectbox("Tolerance", ["2", "5", "10"], index=0, key="tolerance")

    uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])

    temp_path = None
    if uploaded_file is not None:
        # write to a temporary file so that existing reconcile() can read from path
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
        tmp.write(uploaded_file.getbuffer())
        tmp.flush()
        tmp.close()
        temp_path = tmp.name
        st.success("File uploaded successfully.")

    if st.button("Run Reconciliation"):
        if temp_path is None:
            st.error("Please upload an Excel file first.")
        elif taxpayer.strip() == "" or period.strip() == "":
            st.error("Taxpayer GSTIN and Reconciliation Period are required.")
        else:
            run_reconciliation(temp_path, taxpayer, period, tolerance)


def run_reconciliation(file_path, taxpayer, period, tolerance):
    progress = st.progress(0)
    log = make_logger()

    def update_progress(val):
        try:
            progress.progress(int(val))
        except Exception:
            pass

    try:
        output_file, df = reconcile(
            file_path,
            taxpayer,
            period,
            tolerance,
            st.session_state.username,
            progress_callback=update_progress,
            log_callback=log,
        )
        st.success(f"Reconciliation completed. Output saved to: {output_file}")
        # allow download of the generated output
        with open(output_file, "rb") as f:
            st.download_button(
                "Download Output",
                f,
                file_name=os.path.basename(output_file),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        # keep latest results in session state for report generation
        st.session_state.last_output = output_file
        st.session_state.last_reco_df = df

        if df is not None and not df.empty:
            if st.button("Generate Correction Report"):
                out_dir = os.path.dirname(output_file)
                try:
                    path, dfs = generate_correction_report(df, out_dir, log=log)
                    st.success(f"Correction report saved: {path}")
                    with open(path, "rb") as f:
                        st.download_button(
                            "Download Correction Report",
                            f,
                            file_name=os.path.basename(path),
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        )
                except Exception as ex:
                    st.error(f"Failed to generate correction report: {ex}")
    except Exception as exc:
        st.error(str(exc))


def admin_panel():
    st.sidebar.subheader("Admin: User Management")
    users = get_all_users()
    st.sidebar.table(pd.DataFrame(users, columns=["Username", "Role"]))

    new_user = st.sidebar.text_input("New Username", key="new_user")
    new_pwd = st.sidebar.text_input("New Password", type="password", key="new_pwd")
    new_role = st.sidebar.selectbox("Role", ["USER", "ADMIN"], key="new_role")
    if st.sidebar.button("Create User"):
        if new_user and new_pwd:
            if create_user(new_user, new_pwd, new_role):
                st.sidebar.success("User created. Refresh the page to see updates.")
            else:
                st.sidebar.error("Failed to create user (maybe already exists?).")
        else:
            st.sidebar.error("Username and password required.")

    del_user = st.sidebar.text_input("Username to delete", key="del_user")
    if st.sidebar.button("Delete User"):
        if del_user:
            delete_user(del_user)
            st.sidebar.success("Delete command issued. Refresh to update list.")
        else:
            st.sidebar.error("Enter a username to delete.")


if __name__ == "__main__":
    main()
