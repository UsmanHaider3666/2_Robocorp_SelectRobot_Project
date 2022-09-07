from RPA.Browser.Selenium import Selenium
from RPA.HTTP import HTTP
from RPA.Excel.Files import Files
import pandas as pd
from RPA.PDF import PDF
from RPA.Archive import Archive


class SelectRobot:

    def __init__(self):
        self.readfile = None
        self.browse = Selenium()
        self.url = "https://robotsparebinindustries.com/"
        self.http = HTTP()
        self.excel = Files()
        self.pdf = PDF()
        self.zip = Archive()

    def open_browser(self):
        self.browse.open_available_browser(self.url, maximized=True, )
        self.browse.press_keys('//*[@id="root"]/header/div/ul/li[2]/a', "RETURN")
        self.browse.click_button_when_visible('//*[@id="root"]/div/div[2]/div/div/div/div/div/button[1]')
        self.browse.auto_close = False

    def download_the_order_file(self):
        self.http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)

    def convert_to_excel(self):
        self.readfile = pd.read_csv("/home/usman/Python-RPA/2_Robocorp_Select_Robot_Project/orders.csv")
        self.readfile.to_excel("/home/usman/Python-RPA/2_Robocorp_Select_Robot_Project/orders.xlsx")

    def build_and_order_your_robot(self):
        self.excel.open_workbook("orders.xlsx")
        orders_data = self.excel.read_worksheet_as_table(header=True)
        self.excel.close_workbook()
        for data in orders_data:
            while True:
                print(f'Trying for {data["Body"]} {data["Order number"]}')
                self.browse.select_from_list_by_value('//*[@id="head"]', str(data["Head"]))
                self.browse.click_element_when_visible(
                    f'//*[@id="root"]/div/div[1]/div/div[1]/form/div[2]/div/div[{data["Body"]}]/label')
                self.browse.input_text('//*[@placeholder="Enter the part number for the legs"]', str(data["Legs"]))
                self.browse.input_text('//*[@id="address"]', str(data["Address"]))
                self.browse.click_button_when_visible('//*[@id="preview"]')
                self.browse.click_button_when_visible('//*[@id="order"]')
                try:
                    self.browse.get_element_attribute('//*[@class="alert alert-danger"]', attribute="outerHTML")
                    print(f'Not worked {data["Body"]} {data["Order number"]}')
                    self.browse.reload_page()
                    self.browse.click_button_when_visible('//*[@class="btn btn-dark"]')
                except:
                    receipt = self.browse.get_element_attribute('//*[@id="receipt"]', attribute="outerHTML")
                    self.pdf.html_to_pdf(receipt, f"output/receipt{data['Order number']}.pdf")
                    self.browse.screenshot('//*[@id="robot-preview-image"]', f"screenshot{data['Order number']}.png")
                    self.pdf.add_watermark_image_to_pdf(f"screenshot{data['Order number']}.png",
                                                        f"output/receipt{data['Order number']}.pdf",
                                                        f"output/receipt{data['Order number']}.pdf")
                    self.browse.click_button_when_visible('//*[@id="order-another"]')
                    self.browse.click_button_when_visible('//*[@id="root"]/div/div[2]/div/div/div/div/div/button[1]')
                    print(f'Done with {data["Body"]} {data["Order number"]}')
                    break

    def make_zip(self):
        self.zip.archive_folder_with_zip("/home/usman/Python-RPA/2_Robocorp_Select_Robot_Project/output", "orders.zip", True, )


if __name__ == '__main__':
    res = SelectRobot()
    res.open_browser()
    try:
        res.download_the_order_file()
        res.convert_to_excel()
        res.build_and_order_your_robot()
        res.make_zip()
    finally:
        res.browse.close_all_browsers()
