import xlwings as xw
import unittest

import lusid
from lusid.utilities.api_client_builder import ApiClientBuilder
from utilities.credentials_source import CredentialsSource
from utilities.data_source import DataSource


class Scopes(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        api_client = ApiClientBuilder().build(CredentialsSource.secrets_path())

        cls.scopes_api = lusid.ScopesApi(api_client)
        cls.portfolios_api = lusid.PortfoliosApi(api_client)

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

        # get scopes to validate against through python sdk
        response = self.scopes_api.list_scopes()
        validation_scopes = []
        for scope_i in response.values[:9]:
            validation_scopes.append(scope_i.scope)

        # Assert that the scopes from Excel === scopes from LUSID
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

        scope_location = "F17"

        # Step 2: get a list of portfolios
        # get validation
        scope = sht.range(scope_location).value  # Required
        self.assertTrue(scope, msg="Error loading scope from location: " + scope_location)

        effective_at = sht.range("EffectiveDate").value
        as_at = sht.range("AsAtDate").value

        portfolios_validation = []
        portfolios_response_validation = self.portfolios_api.list_portfolios_for_scope(scope)
        for data in portfolios_response_validation.values:
            portfolios_validation.append(
                {
                    "Type": data.type,
                    "ID Scope": data.id.scope,
                    "ID Code": data.id.code,
                    "DisplayName": data.display_name,
                    # "Description": data.description
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
                    # "Description": entry[4]
                }
            )

        portfolios_excel = sorted(portfolios_excel, key=lambda k: k['ID Code'])
        portfolios_validation = sorted(portfolios_validation, key=lambda k: k['ID Code'])

        # Assert that data from excel == data from python SDK
        self.assertEqual(portfolios_excel, portfolios_validation)
        self.assertEqual(portfolios_excel, portfolios_validation)
