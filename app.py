import os
import sys
import subprocess
import pandas as pd
import streamlit as st
from datetime import date, datetime, timedelta
import matplotlib.pyplot as plt
import calendar

# --- Streamlit config ---
st.set_page_config(page_title="CMS Excel", layout="wide")

# --- Auto-install dependency ---
try:
    import openpyxl
except ModuleNotFoundError:
    subprocess.check_call([sys.executable, "-m", "pip", "install", "openpyxl"])
    import openpyxl

# --- Excel setup ---
FILE_PATH = "cms_data.xlsx"

def init_excel():
    if not os.path.exists(FILE_PATH):
        with pd.ExcelWriter(FILE_PATH, engine='openpyxl') as writer:
            pd.DataFrame(columns=["client_id","full_name","phone","email","address","notes"]).to_excel(writer, sheet_name="clients", index=False)
            pd.DataFrame(columns=["order_id","client_id","service_type","weight_count","pickup_date","due_date","status","special_instructions","delivery_fee","total_fee"]).to_excel(writer, sheet_name="orders", index=False)
            pd.DataFrame(columns=["payment_id","order_id","amount_paid","payment_date","payment_method","payment_status","notes"]).to_excel(writer, sheet_name="payments", index=False)
            pd.DataFrame(columns=["expense_id","date_incurred","category","description","amount","fixed_variable","notes"]).to_excel(writer, sheet_name="costs", index=False)

def load_data(sheet):
    try:
        return pd.read_excel(FILE_PATH, sheet_name=sheet)
    except Exception:
        return pd.DataFrame()

def save_data(sheet, df):
    with pd.ExcelWriter(FILE_PATH, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
        df.to_excel(writer, sheet_name=sheet, index=False)

def add_row(sheet, new_row):
    df = load_data(sheet)
    df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
    save_data(sheet, df)

def next_id(df, id_col):
    return int(df[id_col].max()) + 1 if not df.empty else 1

def sql_date(dt):
    return dt.strftime("%Y-%m-%d") if isinstance(dt, (date, datetime)) else str(dt)

def calculate_fee(service, weight, delivery):
    rates = {"WDF":500,"WDI":700,"Iron Only":200,"Bedding":1200}
    rate = next((v for k,v in rates.items() if service.startswith(k)), 500)
    return rate * (weight or 0) + (delivery or 0)

# --- Style ---
st.markdown("""
    <style>
    body, .stApp {
        background-color: #f5f7fa;
        color: #0d47a1;
    }
    .stSidebar {
        background-color: #e3f2fd !important;
    }
    h1, h2, h3, .stMetricLabel {
        color: #0d47a1 !important;
    }
    .stButton>button {
        background-color: #1976d2 !important;
        color: white !important;
        border-radius: 8px !important;
    }
    .stButton>button:hover {
        background-color: #0d47a1 !important;
    }
    </style>
""", unsafe_allow_html=True)

# --- Init Excel ---
init_excel()

# --- Sidebar ---
st.sidebar.title("📘 CMS Menu")
page = st.sidebar.radio("Navigate to", ["Overview","Clients","Client Profile","Orders","Payments & Costs","Calendar","Dashboard"])

# --- Overview ---
if page == "Overview":
    st.title("📊 Wash & Wear CMS Overview")

    clients = load_data("clients")
    orders = load_data("orders")
    payments = load_data("payments")
    costs = load_data("costs")

    total_revenue = payments["amount_paid"].sum() if not payments.empty else 0
    total_costs = costs["amount"].sum() if not costs.empty else 0
    total_profit = total_revenue - total_costs
    pending_orders = len(orders[orders["status"] != "Completed"]) if not orders.empty else 0
    completed_orders = len(orders[orders["status"] == "Completed"]) if not orders.empty else 0
    outstanding_balance = (orders["total_fee"].sum() - total_revenue) if not orders.empty else 0

    # --- Top metrics ---
    col1, col2, col3 = st.columns(3)
    col1.metric("Clients", len(clients))
    col2.metric("Orders", len(orders))
    col3.metric("Total Revenue", f"{total_revenue:,.2f}")

    col4, col5, col6 = st.columns(3)
    col4.metric("Pending Orders", pending_orders)
    col5.metric("Completed Orders", completed_orders)
    col6.metric("Outstanding Balance", f"{outstanding_balance:,.2f}")

    st.divider()

    # --- Financial Summary Chart ---
    if total_revenue > 0 or total_costs > 0:
        st.subheader("💰 Financial Summary")
        fig, ax = plt.subplots()
        ax.bar(["Revenue", "Costs", "Profit"], [total_revenue, total_costs, total_profit],
               color=["#4caf50", "#f44336", "#1976d2"])
        ax.set_title("Revenue vs Costs")
        ax.set_ylabel("Amount (FCFA)")
        st.pyplot(fig)
    else:
        st.info("No financial data available yet.")

    # --- Top Clients ---
    if not payments.empty and not orders.empty and not clients.empty:
        merged = payments.merge(orders, on="order_id").merge(clients, on="client_id")
        top_clients = merged.groupby("full_name")["amount_paid"].sum().nlargest(5)
        st.subheader("💎 Top 5 Clients")
        st.bar_chart(top_clients)
    else:
        st.info("Not enough data to display top clients.")

    # --- Upcoming Deliveries ---
    if not orders.empty:
        orders["due_date"] = pd.to_datetime(orders["due_date"], errors="coerce").dt.date
        upcoming = orders[orders["due_date"] <= date.today() + timedelta(days=7)]
        st.subheader("📅 Upcoming Deliveries (Next 7 Days)")
        if not upcoming.empty:
            st.dataframe(upcoming[["order_id","service_type","status","due_date","total_fee"]],
                         use_container_width=True)
        else:
            st.info("No upcoming deliveries in the next 7 days.")

    # --- Service Distribution ---
    if not orders.empty:
        st.subheader("🧺 Service Distribution")
        service_count = orders["service_type"].value_counts()
        st.bar_chart(service_count)

# --- Clients ---
elif page == "Clients":
    st.header("👥 Manage Clients")
    df = load_data("clients")
    
    with st.form("add_client"):
        name = st.text_input("Full Name")
        phone = st.text_input("Phone")
        email = st.text_input("Email")
        address = st.text_area("Address")
        notes = st.text_area("Notes")
        if st.form_submit_button("Add Client"):
            cid = next_id(df, "client_id")
            add_row("clients", {"client_id":cid,"full_name":name,"phone":phone,"email":email,"address":address,"notes":notes})
            st.success("✅ Client added successfully!")
            df = load_data("clients")

    st.subheader("Existing Clients")
    st.dataframe(df, use_container_width=True)
    
    if not df.empty:
        st.subheader("❌ Delete Client")
        client_to_delete = st.selectbox("Select Client to Delete", df["full_name"])
        confirm = st.checkbox(f"Confirm deletion of {client_to_delete}")
        if st.button("Delete Client") and confirm:
            client_id = df[df["full_name"]==client_to_delete]["client_id"].values[0]
            df = df[df["client_id"] != client_id]
            save_data("clients", df)
            orders = load_data("orders")
            payments = load_data("payments")
            if not orders.empty:
                client_orders = orders[orders["client_id"]==client_id]
                orders = orders[orders["client_id"]!=client_id]
                save_data("orders", orders)
                if not payments.empty and not client_orders.empty:
                    payments = payments[~payments["order_id"].isin(client_orders["order_id"])]
                    save_data("payments", payments)
            st.success(f"✅ Client '{client_to_delete}' and associated orders/payments deleted!")
            df = load_data("clients")

# --- Client Profile ---
elif page == "Client Profile":
    st.header("👤 Client Profile & History")

    clients = load_data("clients")
    orders = load_data("orders")
    payments = load_data("payments")

    if clients.empty:
        st.info("No clients available. Add clients first.")
    else:
        client_name = st.selectbox("Select Client", clients["full_name"])
        client_data = clients[clients["full_name"] == client_name].iloc[0]

        st.subheader(f"Contact Info for {client_name}")
        st.write({
            "Phone": client_data["phone"],
            "Email": client_data["email"],
            "Address": client_data["address"],
            "Notes": client_data["notes"]
        })

        client_orders = orders[orders["client_id"] == client_data["client_id"]]
        client_payments = payments.merge(client_orders[["order_id"]], on="order_id", how="right")

        total_spent = client_payments["amount_paid"].sum() if not client_payments.empty else 0
        total_due = client_orders["total_fee"].sum() - total_spent if not client_orders.empty else 0

        st.metric("Total Spent", f"{total_spent:,.2f}")
        st.metric("Outstanding Balance", f"{total_due:,.2f}")

        st.subheader("📦 Orders")
        st.dataframe(client_orders, use_container_width=True)

        st.subheader("💰 Payments")
        st.dataframe(client_payments, use_container_width=True)

        if st.button("📩 Message Client"):
            st.info("Feature to send email or WhatsApp reminders can be implemented here.")


# --- Orders ---
elif page == "Orders":
    st.header("📦 Orders")

    clients = load_data("clients")
    orders = load_data("orders")

    # --- Add Order Form ---
    with st.form("add_order"):
        client = st.selectbox("Client", clients["full_name"] if not clients.empty else [])
        service = st.selectbox("Service", ["WDF (Wash, Dry, Fold)", "WDI (Wash, Dry, Iron)", "Iron Only", "Bedding"])
        weight = st.number_input("Weight (kg)", min_value=0.0)
        pickup = st.date_input("Pickup Date", value=date.today())
        due = st.date_input("Due Date", value=date.today() + timedelta(days=2))
        delivery = st.number_input("Delivery Fee", min_value=0.0)
        total = calculate_fee(service, weight, delivery)
        status = st.selectbox("Status", ["Scheduled Pickup", "Processing", "Ready", "Completed"])
        note = st.text_area("Notes")

        if st.form_submit_button("Add Order"):
            if clients.empty:
                st.error("⚠️ Please add clients first.")
            else:
                cid = int(clients.loc[clients["full_name"] == client, "client_id"].values[0])
                oid = next_id(orders, "order_id")

                add_row("orders", {
                    "order_id": oid,
                    "client_id": cid,
                    "service_type": service,
                    "weight_count": weight,
                    "pickup_date": sql_date(pickup),
                    "due_date": sql_date(due),
                    "status": status,
                    "special_instructions": note,
                    "delivery_fee": delivery,
                    "total_fee": total
                })

                st.success(f"✅ Order added for {client}!")

    st.subheader("Existing Orders")

    if not orders.empty and not clients.empty:
        # Merge client names for better readability
        merged_orders = orders.merge(clients[["client_id", "full_name"]], on="client_id", how="left")
        st.dataframe(merged_orders, use_container_width=True)

        # --- Update Order Status ---
        st.subheader("🔄 Update Order Status")
        client_choice = st.selectbox("Select Client", merged_orders["full_name"].unique())
        client_orders = merged_orders[merged_orders["full_name"] == client_choice]

        order_choice = st.selectbox("Select Order", client_orders["order_id"].astype(str) + " - " + client_orders["service_type"])
        selected_order_id = int(order_choice.split(" - ")[0])
        selected_order = client_orders[client_orders["order_id"] == selected_order_id].iloc[0]

        st.write(f"**Current Status:** {selected_order['status']}")
        new_status = st.selectbox(
            "New Status",
            ["Scheduled Pickup", "Processing", "Ready", "Completed"],
            index=["Scheduled Pickup", "Processing", "Ready", "Completed"].index(selected_order["status"])
            if selected_order["status"] in ["Scheduled Pickup", "Processing", "Ready", "Completed"]
            else 0
        )

        if st.button("✅ Update Status"):
            orders.loc[orders["order_id"] == selected_order_id, "status"] = new_status
            save_data("orders", orders)
            st.success(f"Order for {client_choice} updated to '{new_status}'!")

        # --- Delete Order ---
        st.subheader("❌ Delete Order")
        del_client = st.selectbox("Select Client to Delete Order", merged_orders["full_name"].unique(), key="del_client")
        del_client_orders = merged_orders[merged_orders["full_name"] == del_client]

        del_order_choice = st.selectbox("Select Order to Delete", del_client_orders["order_id"].astype(str) + " - " + del_client_orders["service_type"])
        del_order_id = int(del_order_choice.split(" - ")[0])

        confirm_del = st.checkbox(f"Confirm deletion of {del_order_choice}")
        if st.button("🗑️ Delete Selected Order") and confirm_del:
            # Delete order and linked payments
            orders = orders[orders["order_id"] != del_order_id]
            save_data("orders", orders)

            payments = load_data("payments")
            payments = payments[payments["order_id"] != del_order_id]
            save_data("payments", payments)

            st.success(f"✅ Order {del_order_choice} for {del_client} deleted successfully!")
    else:
        st.info("No orders available yet. Add orders to get started.")


# --- Payments & Costs ---
elif page == "Payments & Costs":
    st.header("💰 Payments & Costs")

    orders = load_data("orders")
    payments = load_data("payments")
    costs = load_data("costs")

    # --- Add Payment ---
    with st.form("add_payment"):
        order_choice = st.selectbox("Order ID", orders["order_id"] if not orders.empty else [])
        amt = st.number_input("Amount", min_value=0.0)
        pay_date = st.date_input("Payment Date", value=date.today())
        method = st.selectbox("Method", ["Cash", "Mobile Money", "Bank Transfer"])
        status = st.selectbox("Status", ["Paid", "Partial", "Unpaid"])
        notes = st.text_area("Notes")

        if st.form_submit_button("Add Payment"):
            pid = next_id(payments, "payment_id")
            add_row("payments", {
                "payment_id": pid,
                "order_id": order_choice,
                "amount_paid": amt,
                "payment_date": sql_date(pay_date),
                "payment_method": method,
                "payment_status": status,
                "notes": notes
            })
            st.success("✅ Payment recorded!")

    # --- Delete Payment ---
    if not payments.empty:
        st.subheader("❌ Delete Payment")
        payment_to_delete = st.selectbox("Select Payment ID to Delete", payments["payment_id"])
        confirm_payment = st.checkbox(f"Confirm deletion of Payment ID {payment_to_delete}")

        if st.button("Delete Payment") and confirm_payment:
            payments = payments[payments["payment_id"] != payment_to_delete]
            save_data("payments", payments)
            st.success(f"✅ Payment {payment_to_delete} deleted!")

    # --- Add Cost ---
    with st.form("add_cost"):
        d = st.date_input("Date Incurred", value=date.today())
        cat = st.selectbox(
        "Category",
        ["Supplies", "Bills/Rents", "Maintenance", "Others"]
                            )
        desc = st.text_input("Description")
        amt = st.number_input("Amount", min_value=0.0)
        fv = st.selectbox("Type", ["Fixed", "Variable"])
        notes = st.text_area("Notes")

        if st.form_submit_button("Add Cost"):
            eid = next_id(costs, "expense_id")
            add_row("costs", {
                "expense_id": eid,
                "date_incurred": sql_date(d),
                "category": cat,
                "description": desc,
                "amount": amt,
                "fixed_variable": fv,
                "notes": notes
            })
            st.success("✅ Cost added!")

    st.subheader("Payments")
    st.dataframe(load_data("payments"), use_container_width=True)

    st.subheader("Costs")
    st.dataframe(load_data("costs"), use_container_width=True)

# --- Calendar ---
elif page == "Calendar":
    st.header("🗓️ Enhanced Delivery Calendar")

    orders = load_data("orders")
    clients = load_data("clients")

    if not orders.empty and not clients.empty:
        # Merge client names into orders
        orders = orders.merge(clients[["client_id", "full_name"]], on="client_id", how="left")

        if orders.empty:
            st.info("No orders yet.")
        else:
            # --- Initialize session state ---
            if "cal_year" not in st.session_state:
                st.session_state["cal_year"] = date.today().year
            if "cal_month" not in st.session_state:
                st.session_state["cal_month"] = date.today().month
            if "selected_date" not in st.session_state:
                st.session_state["selected_date"] = None

            # --- Month navigation ---
            col1, col2, col3 = st.columns([1, 2, 1])
            if col1.button("⟵"):
                st.session_state["cal_month"] -= 1
            col2.markdown(f"### {calendar.month_name[st.session_state['cal_month']]} {st.session_state['cal_year']}")
            if col3.button("⟶"):
                st.session_state["cal_month"] += 1

            # Handle month overflow
            if st.session_state["cal_month"] < 1:
                st.session_state["cal_month"] = 12
                st.session_state["cal_year"] -= 1
            if st.session_state["cal_month"] > 12:
                st.session_state["cal_month"] = 1
                st.session_state["cal_year"] += 1

            # --- Calendar rendering ---
            orders["due_date"] = pd.to_datetime(orders["due_date"], errors="coerce").dt.date
            today = date.today()
            cal = calendar.Calendar(firstweekday=0)
            month_days = cal.monthdayscalendar(st.session_state["cal_year"], st.session_state["cal_month"])

            for week in month_days:
                cols = st.columns(7)
                for i, day in enumerate(week):
                    if day == 0:
                        cols[i].markdown(
                            "<div style='background:#eceff1;height:80px;border-radius:6px'></div>",
                            unsafe_allow_html=True
                        )
                    else:
                        this_date = date(st.session_state["cal_year"], st.session_state["cal_month"], day)
                        day_orders = orders[orders["due_date"] == this_date]

                        # --- Color logic based on order status ---
                        if not day_orders.empty:
                            if all(day_orders["status"] == "Completed"):
                                color = "#4caf50"  # Green
                            elif any(day_orders["status"].isin(["Scheduled Pickup", "Processing"])):
                                color = "#ffeb3b"  # Yellow
                            elif this_date < today and any(day_orders["status"] != "Completed"):
                                color = "#f44336"  # Red
                            else:
                                color = "#90caf9"  # Blue
                        else:
                            color = "#eceff1"  # Gray (no orders)

                        # --- Calendar day button ---
                        if cols[i].button(f"{day}\n📦 {len(day_orders)}", key=f"cal-{this_date}"):
                            st.session_state["selected_date"] = this_date

                        cols[i].markdown(
                            f"<div style='background:{color};height:40px;border-radius:6px;margin-top:2px'></div>",
                            unsafe_allow_html=True
                        )

            # --- Display orders for selected date ---
            if st.session_state["selected_date"]:
                filtered = orders[orders["due_date"] == st.session_state["selected_date"]]
                with st.expander(
                    f"📋 Orders for {st.session_state['selected_date']} ({len(filtered)})", expanded=True
                ):
                    if filtered.empty:
                        st.info("No orders for this date.")
                    else:
                        st.dataframe(
                            filtered[
                                ["order_id", "full_name", "service_type", "status", "total_fee", "special_instructions"]
                            ],
                            use_container_width=True
                        )

            # --- Show all upcoming deliveries ---
            if st.button("📅 Show All Upcoming Deliveries"):
                upcoming = orders[orders["due_date"] >= today]
                with st.expander("📅 Upcoming Deliveries", expanded=True):
                    st.dataframe(
                        upcoming[
                            ["due_date", "order_id", "full_name", "service_type", "status", "total_fee"]
                        ],
                        use_container_width=True
                    )

    else:
        st.info("Please add clients and orders first to view the calendar.")


# --- Dashboard ---
elif page == "Dashboard":
    st.header("📈 Business Performance Dashboard")

    payments = load_data("payments")
    costs = load_data("costs")

    # --- Monthly Revenue Trend ---
    if not payments.empty:
        payments["payment_date"] = pd.to_datetime(payments["payment_date"], errors="coerce")
        monthly_revenue = payments.groupby(payments["payment_date"].dt.to_period("M"))["amount_paid"].sum().reset_index()
        monthly_revenue["payment_date"] = monthly_revenue["payment_date"].dt.to_timestamp()

        st.subheader("📆 Monthly Revenue Trend")
        fig, ax = plt.subplots()
        ax.plot(monthly_revenue["payment_date"], monthly_revenue["amount_paid"], color="#1976d2", marker="o")
        ax.set_title("Revenue Over Time")
        ax.set_ylabel("Amount (FCFA)")
        ax.grid(True, linestyle='--', alpha=0.5)
        st.pyplot(fig)
    else:
        st.info("No payment data available.")

    # --- Monthly Profit Trend ---
    if not payments.empty and not costs.empty:
        payments["payment_date"] = pd.to_datetime(payments["payment_date"], errors="coerce")
        costs["date_incurred"] = pd.to_datetime(costs["date_incurred"], errors="coerce")

        monthly_rev = payments.groupby(payments["payment_date"].dt.to_period("M"))["amount_paid"].sum()
        monthly_cost = costs.groupby(costs["date_incurred"].dt.to_period("M"))["amount"].sum()
        profit_df = pd.DataFrame({
            "Revenue": monthly_rev,
            "Costs": monthly_cost
        }).fillna(0)
        profit_df["Profit"] = profit_df["Revenue"] - profit_df["Costs"]
        profit_df.index = profit_df.index.to_timestamp()

        # --- Calculate Profit Change ---
        if len(profit_df) >= 2:
            this_month_profit = profit_df["Profit"].iloc[-1]
            last_month_profit = profit_df["Profit"].iloc[-2]
            profit_change = ((this_month_profit - last_month_profit) / last_month_profit * 100) if last_month_profit != 0 else 0
        else:
            this_month_profit = profit_df["Profit"].iloc[-1] if len(profit_df) == 1 else 0
            profit_change = 0

        colA, colB, colC = st.columns(3)
        colA.metric("This Month's Profit", f"{this_month_profit:,.2f} FCFA")
        colB.metric("Change vs Last Month", f"{profit_change:.1f}%", "↑" if profit_change > 0 else "↓")
        colC.metric("Total Months Tracked", len(profit_df))

        st.subheader("📊 Revenue vs Cost vs Profit")
        fig2, ax2 = plt.subplots()
        ax2.plot(profit_df.index, profit_df["Revenue"], label="Revenue", color="#4caf50", marker="o")
        ax2.plot(profit_df.index, profit_df["Costs"], label="Costs", color="#f44336", marker="o")
        ax2.plot(profit_df.index, profit_df["Profit"], label="Profit", color="#1976d2", marker="o")
        ax2.set_title("Monthly Financial Performance")
        ax2.legend()
        ax2.grid(True, linestyle='--', alpha=0.5)
        st.pyplot(fig2)

    # --- Expense Breakdown ---
    if not costs.empty:
        st.subheader("💸 Expenses by Category")
        cat_sum = costs.groupby("category")["amount"].sum()
        fig3, ax3 = plt.subplots()
        ax3.pie(cat_sum, labels=cat_sum.index, autopct="%1.1f%%",
                colors=["#64b5f6","#bbdefb","#90caf9","#42a5f5"])
        ax3.set_title("Expense Breakdown")
        st.pyplot(fig3)
    else:
        st.info("No cost data available.")
