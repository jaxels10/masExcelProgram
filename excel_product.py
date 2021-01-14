from excel_line import excel_line
from summed_months import summed_months
# kan blive et problem at bruge 0'te element, hvis det nogensinde har en tom værdi
class excel_product:
    def __init__(self, row_index=0, product_array=None):
        dummy_array = summed_months()
        dummy_excel_product = excel_line(dummy_array)

        self.row_index = row_index
        self.product_array = product_array
        self.brand = str(product_array[0].brand).capitalize()
        self.code = product_array[0].item_number
        self.description = product_array[0].name
        self.stock_array = []
        self.buy_order_array = []
        self.forecast_array = []
        self.sales_order_array = []
        self.invoiced_array = []

        self.stock_excel_line = dummy_excel_product
        self.buy_order_excel_line = dummy_excel_product

        self.stock_product = dummy_excel_product
        self.buy_order_product = dummy_excel_product
        self.forecast_product = dummy_excel_product
        self.sales_order_product = dummy_excel_product
        self.invoiced_product = dummy_excel_product
        self.total_stock = 0
        self.total_buy_order = 0
        self.total_forecast = 0
        self.total_sales_order = 0
        self.total_invoiced = 0
        self.ptf = 0
        self.fcs_month_1 = 0
        self.fcs = dummy_excel_product




        for product in product_array:
            if "Stock" in product.modul:
                self.stock_array.append(product)
            if "Købsordre" in product.modul:
                self.buy_order_array.append(product)
            if "Forecast" in product.modul:
                self.forecast_array.append(product)
            if "SalgsOrdre" in product.modul:
                self.sales_order_array.append(product)
            if "Invoiced" in product.modul:
                self.invoiced_array.append(product)




        # Udregn stock
        self.sum_1 = 0
        self.sum_2 = 0
        self.sum_3 = 0
        self.sum_4 = 0
        self.sum_5 = 0
        self.sum_6 = 0
        self.sum_7 = 0
        self.sum_8 = 0
        self.sum_9 = 0
        self.sum_10 = 0
        self.sum_11 = 0
        self.sum_12 = 0

        if len(self.stock_array) > 0:
            for line in self.stock_array:
                self.sum_1 = self.sum_1 + line.month_1
                self.sum_2 = self.sum_2 + line.month_2
                self.sum_3 = self.sum_3 + line.month_3
                self.sum_4 = self.sum_4 + line.month_4
                self.sum_5 = self.sum_5 + line.month_5
                self.sum_6 = self.sum_6 + line.month_6
                self.sum_7 = self.sum_7 + line.month_7
                self.sum_8 = self.sum_8 + line.month_8
                self.sum_9 = self.sum_9 + line.month_9
                self.sum_10 = self.sum_10 + line.month_10
                self.sum_11 = self.sum_11 + line.month_11
                self.sum_12 = self.sum_12 + line.month_12

                self.stock_product = summed_months(self.sum_1, self.sum_2, self.sum_3, self.sum_4, self.sum_5,
                                                      self.sum_6, self.sum_7, self.sum_8, self.sum_9, self.sum_10,
                                                      self.sum_11, self.sum_12)

        # Udregn buy_order
        self.sum_1 = 0
        self.sum_2 = 0
        self.sum_3 = 0
        self.sum_4 = 0
        self.sum_5 = 0
        self.sum_6 = 0
        self.sum_7 = 0
        self.sum_8 = 0
        self.sum_9 = 0
        self.sum_10 = 0
        self.sum_11 = 0
        self.sum_12 = 0

        if len(self.buy_order_array) > 0:
            for line in self.buy_order_array:
                self.sum_1 = self.sum_1 + line.month_1
                self.sum_2 = self.sum_2 + line.month_2
                self.sum_3 = self.sum_3 + line.month_3
                self.sum_4 = self.sum_4 + line.month_4
                self.sum_5 = self.sum_5 + line.month_5
                self.sum_6 = self.sum_6 + line.month_6
                self.sum_7 = self.sum_7 + line.month_7
                self.sum_8 = self.sum_8 + line.month_8
                self.sum_9 = self.sum_9 + line.month_9
                self.sum_10 = self.sum_10 + line.month_10
                self.sum_11 = self.sum_11 + line.month_11
                self.sum_12 = self.sum_12 + line.month_12

                self.buy_order_product = summed_months(self.sum_1, self.sum_2, self.sum_3, self.sum_4,
                                                      self.sum_5, self.sum_6, self.sum_7, self.sum_8,
                                                      self.sum_9, self.sum_10, self.sum_11, self.sum_12)

        #Udregn forecast
        self.sum_1 = 0
        self.sum_2 = 0
        self.sum_3 = 0
        self.sum_4 = 0
        self.sum_5 = 0
        self.sum_6 = 0
        self.sum_7 = 0
        self.sum_8 = 0
        self.sum_9 = 0
        self.sum_10 = 0
        self.sum_11 = 0
        self.sum_12 = 0

        if len(self.forecast_array) > 0:
            for line in self.forecast_array:
                self.sum_1 = self.sum_1 + line.month_1
                self.sum_2 = self.sum_2 + line.month_2
                self.sum_3 = self.sum_3 + line.month_3
                self.sum_4 = self.sum_4 + line.month_4
                self.sum_5 = self.sum_5 + line.month_5
                self.sum_6 = self.sum_6 + line.month_6
                self.sum_7 = self.sum_7 + line.month_7
                self.sum_8 = self.sum_8 + line.month_8
                self.sum_9 = self.sum_9 + line.month_9
                self.sum_10 = self.sum_10 + line.month_10
                self.sum_11 = self.sum_11 + line.month_11
                self.sum_12 = self.sum_12 + line.month_12

                self.forecast_product = summed_months(self.sum_1, self.sum_2, self.sum_3, self.sum_4, self.sum_5, self.sum_6, self.sum_7, self.sum_8, self.sum_9, self.sum_10, self.sum_11, self.sum_12)


        #udregn sales_order
        self.sum_1 = 0
        self.sum_2 = 0
        self.sum_3 = 0
        self.sum_4 = 0
        self.sum_5 = 0
        self.sum_6 = 0
        self.sum_7 = 0
        self.sum_8 = 0
        self.sum_9 = 0
        self.sum_10 = 0
        self.sum_11 = 0
        self.sum_12 = 0

        if len(self.sales_order_array) > 0:
            for line in self.sales_order_array:
                self.sum_1 = self.sum_1 + line.month_1
                self.sum_2 = self.sum_2 + line.month_2
                self.sum_3 = self.sum_3 + line.month_3
                self.sum_4 = self.sum_4 + line.month_4
                self.sum_5 = self.sum_5 + line.month_5
                self.sum_6 = self.sum_6 + line.month_6
                self.sum_7 = self.sum_7 + line.month_7
                self.sum_8 = self.sum_8 + line.month_8
                self.sum_9 = self.sum_9 + line.month_9
                self.sum_10 = self.sum_10 + line.month_10
                self.sum_11 = self.sum_11 + line.month_11
                self.sum_12 = self.sum_12 + line.month_12

                self.sales_order_product = summed_months(self.sum_1, self.sum_2, self.sum_3, self.sum_4, self.sum_5,
                                                      self.sum_6, self.sum_7, self.sum_8, self.sum_9, self.sum_10,
                                                      self.sum_11, self.sum_12)

        # udregn invoiced
        self.sum_1 = 0
        self.sum_2 = 0
        self.sum_3 = 0
        self.sum_4 = 0
        self.sum_5 = 0
        self.sum_6 = 0
        self.sum_7 = 0
        self.sum_8 = 0
        self.sum_9 = 0
        self.sum_10 = 0
        self.sum_11 = 0
        self.sum_12 = 0

        if len(self.invoiced_array) > 0:
            for line in self.invoiced_array:
                self.sum_1 = self.sum_1 + line.month_1
                self.sum_2 = self.sum_2 + line.month_2
                self.sum_3 = self.sum_3 + line.month_3
                self.sum_4 = self.sum_4 + line.month_4
                self.sum_5 = self.sum_5 + line.month_5
                self.sum_6 = self.sum_6 + line.month_6
                self.sum_7 = self.sum_7 + line.month_7
                self.sum_8 = self.sum_8 + line.month_8
                self.sum_9 = self.sum_9 + line.month_9
                self.sum_10 = self.sum_10 + line.month_10
                self.sum_11 = self.sum_11 + line.month_11
                self.sum_12 = self.sum_12 + line.month_12

                self.invoiced_product = summed_months(self.sum_1, self.sum_2, self.sum_3, self.sum_4,
                                                         self.sum_5,
                                                         self.sum_6, self.sum_7, self.sum_8, self.sum_9,
                                                         self.sum_10,
                                                         self.sum_11, self.sum_12)

        if self.stock_product is not None:
            self.total_stock = self.stock_product.month_1 + self.stock_product.month_2 + self.stock_product.month_3 + self.stock_product.month_4 + self.stock_product.month_5 + self.stock_product.month_6 + self.stock_product.month_7 + self.stock_product.month_8 + self.stock_product.month_9 + self.stock_product.month_10 + self.stock_product.month_11 + self.stock_product.month_12
        if self.buy_order_product is not None:
            self.total_buy_order = self.buy_order_product.month_1 + self.buy_order_product.month_2 + self.buy_order_product.month_3 + self.buy_order_product.month_4 + self.buy_order_product.month_5 + self.buy_order_product.month_6 + self.buy_order_product.month_7 + self.buy_order_product.month_8 + self.buy_order_product.month_9 + self.buy_order_product.month_10 + self.buy_order_product.month_11 + self.buy_order_product.month_12
        if self.forecast_product is not None:
            self.total_forecast = self.forecast_product.month_1 + self.forecast_product.month_2 + self.forecast_product.month_3 + self.forecast_product.month_4 + self.forecast_product.month_5 + self.forecast_product.month_6 + self.forecast_product.month_7 + self.forecast_product.month_8 + self.forecast_product.month_9 + self.forecast_product.month_10 + self.forecast_product.month_11 + self.forecast_product.month_12
        if self.sales_order_product is not None:
            self.total_sales_order = self.sales_order_product.month_1 + self.sales_order_product.month_2 + self.sales_order_product.month_3 + self.sales_order_product.month_4 + self.sales_order_product.month_5 + self.sales_order_product.month_6 + self.sales_order_product.month_7 + self.sales_order_product.month_8 + self.sales_order_product.month_9 + self.sales_order_product.month_10 + self.sales_order_product.month_11 + self.sales_order_product.month_12
        if self.invoiced_product is not None:
            self.total_invoiced = self.invoiced_product.month_1 + self.invoiced_product.month_2 + self.invoiced_product.month_3 + self.invoiced_product.month_4 + self.invoiced_product.month_5 + self.invoiced_product.month_6 + self.invoiced_product.month_7 + self.invoiced_product.month_8 + self.invoiced_product.month_9 + self.invoiced_product.month_10 + self.invoiced_product.month_11 + self.invoiced_product.month_12

        #self.FC_carryover = summed_months(self.stock_object.month_1 + self.sales_order_product.month_1 + self.invoiced_product.month_1 - self.forecast_product.month_1)
        if self.total_forecast != 0:
            self.ptf = ((self.buy_order_product.month_2 + self.buy_order_product.month_3 + self.buy_order_product.month_4 + self.buy_order_product.month_5 + self.buy_order_product.month_6 + self.buy_order_product.month_7 + self.buy_order_product.month_8 + self.buy_order_product.month_9 + self.buy_order_product.month_10 + self.buy_order_product.month_11 + self.buy_order_product.month_12 + self.stock_product.month_1 + self.sales_order_product.month_1) / self.total_forecast) * 100

        self.fcs_month_1 = self.stock_product.month_1 + self.sales_order_product.month_1 + self.invoiced_product.month_1 - self.forecast_product.month_1
        fcs_month_2 = self.fcs_month_1 + self.buy_order_product.month_2 - self.forecast_product.month_2
        fcs_month_3 = fcs_month_2 + self.buy_order_product.month_3 - self.forecast_product.month_3
        fcs_month_4 = fcs_month_3 + self.buy_order_product.month_4 - self.forecast_product.month_4
        fcs_month_5 = fcs_month_4 + self.buy_order_product.month_5 - self.forecast_product.month_5
        fcs_month_6 = fcs_month_5 + self.buy_order_product.month_6 - self.forecast_product.month_6
        fcs_month_7 = fcs_month_6 + self.buy_order_product.month_7 - self.forecast_product.month_7
        fcs_month_8 = fcs_month_7 + self.buy_order_product.month_8 - self.forecast_product.month_8
        fcs_month_9 = fcs_month_8 + self.buy_order_product.month_9 - self.forecast_product.month_9
        fcs_month_10 = fcs_month_9 + self.buy_order_product.month_10 - self.forecast_product.month_10
        fcs_month_11 = fcs_month_10 + self.buy_order_product.month_11 - self.forecast_product.month_11
        fcs_month_12 = fcs_month_11 + self.buy_order_product.month_12 - self.forecast_product.month_12

        self.fcs = summed_months(self.fcs_month_1, fcs_month_2, fcs_month_3, fcs_month_4, fcs_month_5, fcs_month_6, fcs_month_7, fcs_month_8, fcs_month_9, fcs_month_10, fcs_month_11, fcs_month_12)