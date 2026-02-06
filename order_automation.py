from playwright.sync_api import sync_playwright
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
import os
import pandas as pd

HEADLESS = False  # debugging
#HEADLESS = True   # production

UNIT_SELECTOR_MAP = {
    "CA": None,  # default, no action
    "PK": "a99681bb3f714c538839c7ee3763a9c2",
    "PCS": "a4c94a3af0974ee791dfcac97327e3b9",
    "UN": "a4c94a3af0974ee791dfcac97327e3b9"
}


# =========================
# CONFIGURATION
# =========================

BASE_URL = "https://pzng.dms.nexxadatatech.com"

ROUTES = {
    "login": f"{BASE_URL}/",
    "orders": f"{BASE_URL}/pre-orders",
    "new_order": f"{BASE_URL}/pre-orders/new-order",
}

USERNAME = os.getenv("NEXXA_USERNAME")
PASSWORD = os.getenv("NEXXA_PASSWORD")
#USERNAME = "biodunkpo"
#PASSWORD = "KPO@8910"
ORDER_URL_KEYWORD = "/pre-orders/new-order"
DASHBOARD_URL_KEYWORD = "/dashboard"
LOGIN_FORM_SELECTOR = "input[name='username']"

# Fixed distributor (global truth)
DISTRIBUTOR = "BIODUN BOLARINWA"


# =========================
# HELPERS
# =========================

def log(msg):
    print(f"[INFO] {msg}")


def go_to(page, route_key):
    url = ROUTES.get(route_key)
    if not url:
        raise ValueError(f"Unknown route: {route_key}")

    log(f"Navigating to {route_key}")
    page.goto(url)
    page.wait_for_load_state("networkidle")


# =========================
# AUTHENTICATION
# =========================

def is_logged_in(page) -> bool:
    """
    Determines if user is already logged in by checking:
    - URL
    - Dashboard-specific UI
    """
    if DASHBOARD_URL_KEYWORD in page.url:
        return True

    try:
        # Element that ONLY exists after login (sidebar / menu)
        page.get_by_role("combobox", name="Sales").wait_for(timeout=3000)
        return True
    except PlaywrightTimeoutError:
        return False


def login(page):
    log("Checking login state")

    go_to(page, "login")

    # If already logged in, skip login
    if is_logged_in(page):
        log("Already logged in")
        return

    log("Performing login")

    if not USERNAME or not PASSWORD:
        raise RuntimeError(
            "Missing credentials. Set NEXXA_USERNAME and NEXXA_PASSWORD "
            "as environment variables before running."
        )

    page.fill("input[name='username']", USERNAME)
    page.fill("input[name='password']", PASSWORD)
    page.click("button[type='submit']")

    try:
        # Wait for dashboard redirect or dashboard element
        page.wait_for_url(f"**{DASHBOARD_URL_KEYWORD}**", timeout=8000)
        log("Login redirect detected")

    except PlaywrightTimeoutError:
        log("Login redirect not detected, checking page state")

    # Final verification
    if not is_logged_in(page):
        page.screenshot(path="login_error.png")
        raise RuntimeError("❌ Login failed: Dashboard not detected")

    log("✅ Login successful — Dashboard confirmed")


# =========================
# ORDER BUILDING
# =========================

def load_orders_from_excel(path="orders.xlsx"):
    # ---------- Read Excel ----------
    try:
        df = pd.read_excel(path, sheet_name="Orders")
        print("[INFO] Reading data from sheet: 'Orders'")
    except ValueError:
        df = pd.read_excel(path)
        print("[INFO] 'Orders' sheet not found — using first sheet")

    # ---------- Required columns ----------
    required_cols = {"order_id", "sales_rep", "outlet", "product", "qty"}
    optional_cols = {"unit"}

    missing = required_cols - set(df.columns)
    if missing:
        raise RuntimeError(f"Missing required columns in Excel: {missing}")

    # ---------- Normalize column names ----------
    df.columns = [c.strip().lower() for c in df.columns]

    orders = {}

    for index, row in df.iterrows():
        excel_row = index + 2  # Excel row number

        # ---------- Clean values ----------
        if pd.isna(row["order_id"]):
            raise RuntimeError(f"order_id missing at row {excel_row}")

        order_id = row["order_id"]

        sales_rep = str(row["sales_rep"]).strip()
        outlet = str(row["outlet"]).strip()
        product = str(row["product"]).strip()

        if not sales_rep or not outlet or not product:
            raise RuntimeError(f"Empty field at row {excel_row}")

        if pd.isna(row["qty"]):
            raise RuntimeError(f"Qty missing at row {excel_row}")

        qty = int(float(row["qty"]))

        # ---------- Unit handling ----------
        unit = "CA"
        if "unit" in df.columns and not pd.isna(row.get("unit")):
            unit = str(row["unit"]).strip().upper()

        # Normalize common aliases
        if unit in ["CARTON", "CTN"]:
            unit = "CA"
        elif unit in ["PACK", "PKT"]:
            unit = "PK"
        elif unit in ["PCS", "PIECE", "UNIT", "UN"]:
            unit = "PCS"

        if unit not in ["CA", "PK", "PCS"]:
            raise RuntimeError(f"Invalid unit '{unit}' at row {excel_row}")

        # ---------- Group by order_id ----------
        if order_id not in orders:
            orders[order_id] = {
                "sales_rep": sales_rep,
                "outlet": outlet,
                "products": []
            }
        else:
            # Safety check: same order_id must not mix outlet/sales rep
            if orders[order_id]["sales_rep"] != sales_rep:
                raise RuntimeError(
                    f"Sales rep mismatch for order {order_id} at row {excel_row}"
                )
            if orders[order_id]["outlet"] != outlet:
                raise RuntimeError(
                    f"Outlet mismatch for order {order_id} at row {excel_row}"
                )

        orders[order_id]["products"].append({
            "name": product,
            "qty": qty,
            "unit": unit
        })

    print(f"[INFO] Loaded {len(orders)} orders from Excel")
    #print(f"Order {order_id} | {sales_rep} | {outlet} | Product: {product} | Qty: {qty} | Unit: {unit}")
    print(f"[DEBUG] Loaded order {order_id} with {len(orders[order_id]['products'])} items")

    return list(orders.values())


def open_new_order(page, sales_rep, outlet):
    log("Opening new order")
    go_to(page, "new_order")
    
    log("now in order page")
    
    # Distributor
    page.get_by_role("combobox").filter(has_text="Select").first.click()
    log("clicked distributor box")
    page.get_by_role("option", name=DISTRIBUTOR).click()
    log("set distributor box")

    # Sales Rep
    page.get_by_role("combobox").filter(has_text="Select").nth(1).click()
    page.get_by_label("Select").get_by_text(sales_rep).click()

    # Outlet
    page.get_by_role("combobox").filter(has_text="Select").click()
    page.get_by_label("Select").get_by_text(outlet).click()

    # Open product modal
    page.get_by_role("button", name="Add Product").click()
    

def add_product(page, product_name, quantity, unit="CA"):
    unit = unit.strip().upper()
    log(f"Adding product: {product_name} | {quantity} | {unit}")

    # 1 Search product
    search_box = page.get_by_role("textbox", name="Search product")
    search_box.click()
    search_box.press("ControlOrMeta+a")
    search_box.fill(product_name)
    log("Product searched")

    # 2 Enter quantity
    qty_box = page.get_by_role("textbox", name="1")
    qty_box.click()
    qty_box.press("ControlOrMeta+a")
    qty_box.fill(str(quantity))
    log(f"Quantity entered: {quantity}")
    
    # 3 Change unit IF NOT carton
    if unit != "CA":
        selector_value = UNIT_SELECTOR_MAP.get(unit)

        if not selector_value:
            raise ValueError(f"Unsupported unit: {unit}")

        page.locator("select").nth(4).select_option(selector_value)
        log(f"Unit switched to {unit}")

    # 4 Add product
    page.get_by_role("button", name="Add", exact=True).click()
    log("Product added")


def finalize_products(page):
    page.get_by_role("button", name="cancel", exact=True).click()
    page.get_by_role("button", name="Proceed").click()
    log("Products finalized")


# =========================
# SUBMIT & FULFILL
# =========================
def is_submitted(page) -> bool:
    """
    Determines if user is already logged in by checking:
    - URL
    - Dashboard-specific UI
    """
    if ORDER_URL_KEYWORD not in page.url:
        return True

    try:
        # Element that ONLY exists after submit (sidebar / menu)
        page.get_by_role("button", name="New Order").wait_for(timeout=5000)
        return True
    except PlaywrightTimeoutError:
        return False
    
def submit_and_fulfill(page):
    log("Submitting order")

    page.get_by_role("checkbox", name="check-item").first.check()
    page.get_by_role("checkbox", name="check-item").nth(3).check()
    if is_submitted(page):
        log("submitted")
    else:
        page.get_by_role("button", name="Submit").click()
        log("auto submitted")
    order_ref=1
    '''page.wait_for_load_state("networkidle")

    order_ref = page.get_by_role("link").first.inner_text()
    log(f"Order Reference: {order_ref}")

    page.get_by_role("link", name=order_ref).click()

    # Uncomment when needed
    page.get_by_role("button", name="approve Fulfill").click()
    page.get_by_role("button", name="Fulfill", exact=True).click()'''

    log("Order fulfilled")
    return order_ref


# =========================
# ORCHESTRATOR
# =========================


def run_order(page, order):
    open_new_order(page, order["sales_rep"], order["outlet"])

    for item in order["products"]:
        # pass the unit too
        add_product(page, item["name"], item["qty"], item.get("unit", "CA"))

    finalize_products(page)
    return submit_and_fulfill(page)

# =========================
# MAIN
# =========================

def main():
    orders = load_orders_from_excel(r"C:\Users\User\Documents\Orders.xlsx")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=HEADLESS)
        context = browser.new_context()
        page = context.new_page()

        login(page)

        for order in orders:
            ref = run_order(page, order)
            log(f"ORDER COMPLETED: {ref}")

        input("Press ENTER to close browser...")
        browser.close()


if __name__ == "__main__":
    main()
