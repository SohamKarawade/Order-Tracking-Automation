import subprocess
import os
import customtkinter as ctk
from PIL import Image
import win32com.client
import subprocess

ctk.set_appearance_mode("dark")

app = ctk.CTk()
app.geometry("1100x700")
app.title("Shopify Order Automation")
app.resizable(False, False)

# =====================================
# BACKGROUND IMAGE
# =====================================

bg_img = ctk.CTkImage(
    light_image=Image.open("C:\\Users\\Midaashna\\OneDrive\\Desktop\\mail automation\\dist\\assets\\bgm.png"),
    dark_image=Image.open("C:\\Users\\Midaashna\\OneDrive\\Desktop\\mail automation\\dist\\assets\\bgm.png" \
    ""),
    size=(1100,700)
)

bg_label = ctk.CTkLabel(
    app,
    image=bg_img,
    text=""
)

bg_label.place(relx=0, rely=0, relwidth=1, relheight=1)

# =====================================
# BUTTONS
# =====================================

def send_report():

    btn2.configure(text="Generating............")
    app.update()

    # Generate latest report
    subprocess.run([
        "python",
        "project.py"
    ])
    outlook = win32com.client.Dispatch(
        "Outlook.Application"
    )

    mail = outlook.CreateItem(0)

    mail.To = "amit@ritvaa.in"

    mail.Subject = (
        "Shopify Order Tracking Report"
    )

    mail.Body = """
Hello Amit Sir,

Please find attached the latest Shopify Order Tracking Report.

The report contains:

• Order Summary
• Pending Orders
• Confirmed Orders
• Delivered Orders
• Cancelled Orders
• Replacement Orders

Kindly review and let us know if any action is required.

Regards,
Nayan Anil Sawant
"""

    attachment_path = os.path.abspath(
        "Order_Tracking.xlsx"
    )

    mail.Attachments.Add(
        attachment_path
    )

    mail.Save()

    btn2.configure(text="📤 Send Report")


def open_report():

    btn3.configure(text="Generating............")
    app.update()

    subprocess.run(["python", "project.py"])

    os.startfile("Order_Tracking.xlsx")

    btn3.configure(text="📊 Generate Report")

btn2 = ctk.CTkButton(
    app,
    text="📤 Send Report",
    width=130,
    height=50,
    corner_radius=12,
    fg_color="#04001B",
    hover_color="#400050",
    border_width=1,
    border_color="#400050",
    text_color="white",
    font=("Segoe UI", 13, "bold"),
    command=send_report
)

btn2.place(x=390, y=340)


btn3 = ctk.CTkButton(
    app,
    text="📊 Generate Report",
    width=100,
    height=50,
    corner_radius=12,
    fg_color="#04001B",
    hover_color="#400050",
    border_width=1,
    border_color="#400050",
    text_color="white",
    font=("Segoe UI", 13, "bold"),
    command=open_report
)

btn3.place(x=525, y=340)

app.mainloop()

