import xlwings as xw
import unittest
from datetime import datetime
import pytz

import lusid
from lusid import models
from lusid.utilities.api_client_builder import ApiClientBuilder
from utilities.credentials_source import CredentialsSource
from utilities.data_source import DataSource


def get_date(date, sht):
    date_at = sht.range(date).value
    if date_at:
        date_at = datetime(date_at.year, date_at.month, date_at.day, tzinfo=pytz.utc)
    return date_at


class Scopes(unittest.TestCase):
    # This test validates that the data displayed in "LUSID Excel -Setting up your IBOR template Global Wquity 
    # Fund.xlsx" is correct using the python SDK. 

    @classmethod
    def setUpClass(cls):
        api_client = ApiClientBuilder().build(CredentialsSource.secrets_path())

        cls.scopes_api = lusid.ScopesApi(api_client)
        cls.portfolios_api = lusid.PortfoliosApi(api_client)
        cls.transaction_portfolios_api = lusid.TransactionPortfoliosApi(api_client)

    def test_list_scopes(self):

        # define location of data in excel spreadsheet
        book_name = "LUSID Excel - Setting up your IBOR template Global Equity Fund.xlsx"
        book_path = DataSource.data_path(book_name)
        sheet_name = 'View scopes'
        range_containing_scopes = 'D16:D26'

        # get scopes from excel
        wb = xw.Book(book_path)
        self.assertTrue(wb, msg="Error loading book: " + book_name)

        sht = wb.sheets(sheet_name)
        self.assertTrue(sht, msg="Error loading sheet: " + sheet_name)

        scopes = sht.range(range_containing_scopes).value
        self.assertTrue(scopes, msg="Error loading scopes from range: " + range_containing_scopes)

        excel_scopes = []
        scopes = scopes[0:9]
        for scope in scopes:
            if scope:
                excel_scopes.append(scope)

        # get scopes to validate against
        response = self.scopes_api.list_scopes()
        validation_scopes = []
        for scope_i in response.values[:9]:
            validation_scopes.append(scope_i.scope)

        # Assert that the scopes from Excel == scopes from LUSID
        self.assertEqual(excel_scopes, validation_scopes, msg="Scopes from " + book_name + "do not equal validation "
                                                                                           "set")

    def test_view_portfolios(self):

        # define location of data in excel spreadsheet
        book_name = "LUSID Excel - Setting up your IBOR template Global Equity Fund.xlsx"
        book_path = DataSource.data_path(book_name)
        sheet_name = 'View portfolios'
        wb = xw.Book(book_path)
        self.assertTrue(wb, msg="Error loading book: " + book_name)
        sht = wb.sheets(sheet_name)
        self.assertTrue(sht, msg="Error loading sheet: " + sheet_name)

        # define locations/names of cells containing fields
        scope_location = "F17"
        EffectiveDate = "EffectiveDate"
        AsAtDate = "AsAtDate"

        effective_at = get_date(EffectiveDate, sht)
        as_at = get_date(AsAtDate, sht)

        # Step 2: get a list of portfolios
        # get validation
        scope = sht.range(scope_location).value  # Required
        self.assertTrue(scope, msg="Error loading scope from location: " + scope_location)

        portfolios_validation = []
        portfolios_response_validation = self.portfolios_api.list_portfolios_for_scope(
            scope=scope,
            effective_at=effective_at if effective_at else "",
            as_at=as_at if as_at else ""
        )

        for data in portfolios_response_validation.values:
            portfolios_validation.append(
                {
                    "Type": data.type,
                    "ID Scope": data.id.scope,
                    "ID Code": data.id.code,
                    "DisplayName": data.display_name,
                }
            )

        # get excel
        portfolios_excel = []
        portfolios_response_excel = sht.range("E30", "M39").value
        for entry in portfolios_response_excel:
            if (entry[0] == "") or (entry[0] == None):
                continue
            portfolios_excel.append(
                {
                    "Type": entry[0],
                    "ID Scope": entry[1],
                    "ID Code": entry[2],
                    "DisplayName": entry[3],
                }
            )

        # Sort data by ID
        portfolios_excel = sorted(portfolios_excel, key=lambda k: k["ID Code"])
        portfolios_validation = sorted(portfolios_validation, key=lambda k: k["ID Code"])

        # Assert that data from excel == data from python SDK
        self.assertEqual(portfolios_excel, portfolios_validation)
        self.assertEqual(portfolios_excel, portfolios_validation)

    def test_view_holdings(self):

        # define location of data in excel spreadsheet
        book_name = "LUSID Excel - Setting up your IBOR template Global Equity Fund.xlsx"
        book_path = DataSource.data_path(book_name)
        sheet_name = 'View holdings'
        wb = xw.Book(book_path)
        self.assertTrue(wb, msg="Error loading book: " + book_name)
        sht = wb.sheets(sheet_name)
        self.assertTrue(sht, msg="Error loading sheet: " + sheet_name)

        # define locations/names of cells containing fields
        scope_location = "F17"
        PortfolioCode = "F19"
        EffectiveDate = "F23"
        AsAtDate = "F25"

        effective_at = get_date(EffectiveDate, sht)
        as_at = get_date(AsAtDate, sht)

        # Step 2: get a list of holdings
        # get validation
        scope = sht.range(scope_location).value  # Required
        portfolio_code = sht.range(PortfolioCode).value  # Required
        self.assertTrue(scope, msg="Error loading scope from location: " + scope_location)
        self.assertTrue(portfolio_code, msg="Error loading scope from location: " + PortfolioCode)

        holdings_response = self.transaction_portfolios_api.get_holdings(
            scope=scope,
            code=portfolio_code,
            effective_at=effective_at if effective_at else "",
            as_at=as_at if as_at else ""
        )

        holdings_validation = []
        for data in holdings_response.values:
            holdings_validation.append(
                {
                    "InstrumentUid": data.instrument_uid,
                    "HoldingType": data.holding_type,
                    "Units": data.units,
                    "SettledUnits": data.settled_units,
                    "Cost Amount": data.cost.amount,
                    "Cost Currency": data.cost.currency,
                    "CostPortfolioCcy Amount": data.cost_portfolio_ccy.amount,
                    "CostPortfolioCcy Currency": data.cost_portfolio_ccy.currency,
                }
            )

        # get holding from excel
        holdings_excel = []
        holdings_response_excel = sht.range("E32", "L510").value
        for entry in holdings_response_excel:
            if (entry[0] == "") or (entry[0] == None):
                continue
            holdings_excel.append(
                {
                    "InstrumentUid": entry[0],
                    "HoldingType": entry[1],
                    "Units": entry[2],
                    "SettledUnits": entry[3],
                    "Cost Amount": entry[4],
                    "Cost Currency": entry[5],
                    "CostPortfolioCcy Amount": entry[6],
                    "CostPortfolioCcy Currency": entry[7],
                }
            )

        holdings_excel = sorted(holdings_excel, key=lambda k: k["InstrumentUid"])
        holdings_validation = sorted(holdings_validation, key=lambda k: k["InstrumentUid"])

        # Assert that data from excel == data from python SDK
        self.assertEqual(holdings_excel, holdings_validation)
