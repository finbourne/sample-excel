import xlwings as xw
import unittest

import lusid
from lusid.utilities.api_client_builder import ApiClientBuilder
from utilities.credentials_source import CredentialsSource


class Scopes(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        api_client = ApiClientBuilder().build(CredentialsSource.secrets_path())

        cls.scopes_api = lusid.ScopesApi(api_client)
        cls.portfolios_api = lusid.PortfoliosApi(api_client)



    def test_list_scopes(self):
        # define location of data in excel spreadsheet
        book_name = "LUSID Excel - Setting up your IBOR template Global Equity Fund.xlsx"
        sheet_name = 'View scopes'
        range_containing_scopes = 'D16:D26'

        # get scopes from excel
        wb = xw.Book(book_name)
        self.assertTrue(wb, msg="Error loading book: " + book_name)

        sht = wb.sheets(sheet_name)
        self.assertTrue(sht, msg="Error loading sheet: " + sheet_name)

        scopes = sht.range(range_containing_scopes).value
        self.assertTrue(scopes, msg="Error loading scopes from range: " + range_containing_scopes)

        excel_scopes = []
        for scope in scopes:
            if scope:
                excel_scopes.append(scope)

        # get scopes to validate against through python sdk
        response = self.scopes_api.list_scopes()
        validation_scopes = []
        for scope_i in response.values:
            validation_scopes.append(scope_i.scope)

        # Assert that the scopes from Excel === scopes from LUSID
        self.assertEqual(excel_scopes, validation_scopes, msg="Scopes from " + book_name + "do not equal validation "
                                                                                           "set")