import os
import time
import xlwings as xw
import unittest
from datetime import datetime
import pytz
import json
import pprint
from dateutil import tz

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


def open_excel(wait_time):

    os.environ["fbn-excel-base-api-url"] = "<apiurl>"
    os.environ["fbn-excel-auth-redirect-url"] = "<redirect-url>"
    os.environ["fbn-excel-auth-uri"] = "<auth-uri>"
    os.environ["fbn-excel-auth-client-id"] = "<CLIENT-ID>"
    os.environ["FBN_TOKEN_URL"] = "<token>"
    os.environ["FBN_LUSID_API_URL"] = "<API_URL>"
    os.environ["FBN_CLIENT_ID"] = "<CLIENT_ID>"
    os.environ["FBN_CLIENT_SECRET"] = "<CLIENT_SECRET>"
    os.environ["FBN_USERNAME"] = "<USERNAME>"
    os.environ["FBN_PASSWORD"] = "<PASSWORD>"

    # Requires excel addin to be installed in EXCEL
    os.startfile(r"EXCEL.EXE")

    time.sleep(wait_time)
    return 0


class IborTemplate(unittest.TestCase):
    # This test validates that the data displayed in "LUSID Excel -Setting up your IBOR template Global Equity
    # Fund.xlsx" is correct using the python SDK.

    @classmethod
    def setUpClass(cls):

        cls.wait_time = 15
        open_excel(cls.wait_time)

        api_client = ApiClientBuilder().build(CredentialsSource.secrets_path())

        cls.scopes_api = lusid.ScopesApi(api_client)
        cls.portfolios_api = lusid.PortfoliosApi(api_client)
        cls.transaction_portfolios_api = lusid.TransactionPortfoliosApi(api_client)
        cls.aggregation_api = lusid.AggregationApi(api_client)
        cls.reconciliations_api = lusid.ReconciliationsApi(api_client)
        cls.book_name = "LUSID Excel - Setting up your IBOR template Global Equity Fund.xlsx"

    def get_sheet(self, sheet_name):
        book_path = DataSource.data_path(self.book_name)
        wb = xw.Book(book_path)
        app1 = xw.apps
        app1.active.calculate()
        self.assertTrue(wb, msg="Error loading book: " + self.book_name)
        sht = wb.sheets(sheet_name)
        self.assertTrue(sht, msg="Error loading sheet: " + sheet_name)

        return sht

    def test_list_scopes(self):

        sht = self.get_sheet('View scopes')

        # get parameters from excel
        range_containing_scopes = 'D16:D26'

        # get excel data

        scopes = sht.range(range_containing_scopes).value
        self.assertTrue(scopes, msg="Error loading scopes from range: " + range_containing_scopes)

        # format excel data
        excel_scopes = [scope for scope in scopes[:9] if scope]

        # get validation data
        response = self.scopes_api.list_scopes()
        validation_scopes = [scope_i.scope for scope_i in response.values]

        # Assert excel data is contained within validation
        [self.assertIn(sample, validation_scopes, msg="Scopes from " + self.book_name + "do not equal validation ") for
         sample in excel_scopes]

    def test_view_portfolios(self):

        sht = self.get_sheet('View portfolios')

        # get parameters from excel
        scope_location = "F17"
        EffectiveDate = "EffectiveDate"
        AsAtDate = "AsAtDate"

        scope = sht.range(scope_location).value
        effective_at = get_date(EffectiveDate, sht)
        as_at = get_date(AsAtDate, sht)

        self.assertTrue(scope, msg="Error loading scope from location: " + scope_location)

        # get validation data
        portfolios_response_validation = self.portfolios_api.list_portfolios_for_scope(
            scope=scope,
            effective_at=effective_at if effective_at else "",
            as_at=as_at if as_at else ""
        )

        # format validation data
        portfolios_validation = [
            {
                "Type": data.type,
                "ID Scope": data.id.scope,
                "ID Code": data.id.code,
                "DisplayName": data.display_name,
            }
            for data in portfolios_response_validation.values
        ]

        # get excel data
        portfolios_response_excel = sht.range("E30", "M39").value

        # format excel data
        portfolios_excel = [
            {
                "Type": entry[0],
                "ID Scope": entry[1],
                "ID Code": entry[2],
                "DisplayName": entry[3],
            }
            for entry in portfolios_response_excel if (entry[0] != "" and entry[0] is not None)
        ]

        # Assert excel data is contained within validation
        [self.assertIn(sample, portfolios_validation) for sample in portfolios_excel]

    def test_view_holdings(self):

        sht = self.get_sheet('View holdings')

        # get parameters from excel
        scope_location = "F17"
        PortfolioCode = "F19"
        EffectiveDate = "F23"
        AsAtDate = "F25"

        effective_at = get_date(EffectiveDate, sht)
        as_at = get_date(AsAtDate, sht)
        scope = sht.range(scope_location).value
        portfolio_code = sht.range(PortfolioCode).value

        self.assertTrue(scope, msg="Error loading scope from location: " + scope_location)
        self.assertTrue(portfolio_code, msg="Error loading scope from location: " + PortfolioCode)

        # get validation data
        holdings_response = self.transaction_portfolios_api.get_holdings(
            scope=scope,
            code=portfolio_code,
            effective_at=effective_at if effective_at else "",
            as_at=as_at if as_at else ""
        )

        holdings_validation = [
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
            for data in holdings_response.values
        ]

        # get holding from excel
        holdings_response_excel = sht.range("E32", "L510").value
        holdings_excel = [
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
            for entry in holdings_response_excel if (entry[0] != "" and entry[0] is not None)
        ]

        holdings_excel = sorted(holdings_excel, key=lambda k: k["InstrumentUid"])
        holdings_validation = sorted(holdings_validation, key=lambda k: k["InstrumentUid"])

        # Assert that data from excel == data from python SDK
        [self.assertIn(sample, holdings_validation) for sample in holdings_excel]

    @unittest.skip("'run_valuation' Test not fully implemented")
    def test_run_valuation(self):
        # define location of data in excel spreadsheet
        sht = self.get_sheet('Run valuation')

        # define locations/names of cells containing fields
        scope_location = "F17"
        PortfolioCode = "F19"
        ValuationDate = "F21"  # effective valuation date
        InlineRecipe = "J35"

        effective_at = get_date(ValuationDate, sht)
        scope = sht.range(scope_location).value
        code = sht.range(PortfolioCode).value
        inline_recipe = json.loads(sht.range(InlineRecipe).value)
        recipe_keys = inline_recipe.keys()

        self.assertTrue(scope, msg="Error loading scope from location: " + scope_location)
        self.assertTrue(code, msg="Error loading scope from location: " + PortfolioCode)

        print(f"Scope:          {scope}")
        print(f"Code:           {code}")
        print(f"Effective Date: {effective_at}")

        # get valuation data
        inline_recipe = models.ConfigurationRecipe(
            code='quotes_recipe',
            market=models.MarketContext(
                market_rules=[],
                suppliers=models.MarketContextSuppliers(
                    equity='DataScope'
                ),
                options=models.MarketOptions(
                    default_supplier='DataScope',
                    default_instrument_code_type='LusidInstrumentId',
                    default_scope=scope)
            )
        )

        aggregation_request = models.AggregationRequest(
            inline_recipe=inline_recipe,
            metrics=[
                models.AggregateSpec("Instrument/default/Name", "Value"),
                models.AggregateSpec("Holding/default/PV", "Proportion"),
                models.AggregateSpec("Holding/default/PV", "Sum")
            ],
            group_by=["Instrument/default/Name"],
            effective_at=effective_at
        )

        print(f"agg req:")
        pprint.pprint(aggregation_request)

        response = self.aggregation_api.get_aggregation(scope=scope, code=code, request=aggregation_request)

        for item in response.data:
            print("\t{}\t{}\t{}".format(item["Instrument/default/Name"], item["Proportion(Holding/default/PV)"],
                                        item["Sum(Holding/default/PV)"]))
        print("-------------------------------------------------")
        pprint.pprint(aggregation_request)

    def test_view_transactions(self):
        # define location of data in excel spreadsheet
        sht = self.get_sheet('View transactions')

        # get parameters from excel
        scope_location = "F17"
        PortfolioCode = "F19"
        FromTransactionDate = "F23"
        ToTransactionDate = "F25"
        AsAt = "F27"

        scope = sht.range(scope_location).value
        code = sht.range(PortfolioCode).value
        from_transaction_date = get_date(FromTransactionDate, sht)
        to_transaction_date = get_date(ToTransactionDate, sht)
        as_at = get_date(AsAt, sht)

        self.assertTrue(scope, msg="Error loading scope from location: " + scope_location)
        self.assertTrue(code, msg="Error loading scope from location: " + PortfolioCode)

        # get Excel data
        transaction_response_excel = sht.range("E33", "M51").value
        headers = transaction_response_excel.pop(0)

        # get validation data
        build_data = self.transaction_portfolios_api.build_transactions(
            scope=scope,
            code=code,
            parameters=models.TransactionQueryParameters(
                start_date=from_transaction_date if from_transaction_date else "",
                end_date=to_transaction_date if to_transaction_date else "",
                query_mode='TradeDate',
                show_cancelled_transactions=False
            )
        )

        # format validation data
        transaction_validation = [
            {
                "TransactionId": data.transaction_id,
                "Type": data.type,
                "InstrumentUid": data.instrument_uid,
                "TransactionDate": data.transaction_date,
                "Units": data.units,
            }
            for data in build_data.values
        ]

        # format excel data
        transaction_excel = [
            {
                "TransactionId": data[headers.index('TransactionId')],
                "Type": data[headers.index('Type')],
                "InstrumentUid": data[headers.index('InstrumentUid')],
                "TransactionDate": data[headers.index('TransactionDate')],
                "Units": data[headers.index('Units')],
            }
            for data in transaction_response_excel if (data[0] != "" and data[0] is not None)
        ]

        for transaction in transaction_excel:
            transaction["TransactionDate"] = datetime(transaction["TransactionDate"].year,
                                                      transaction["TransactionDate"].month,
                                                      transaction["TransactionDate"].day, tzinfo=tz.tzutc())

        # perform assertion
        [self.assertIn(sample, transaction_validation) for sample in transaction_excel]

    @unittest.skip("'run_valuation' Test not fully implemented: excel copies headers???")
    def test_Perform_a_reconciliation(self):

        # load sheet
        sht = self.get_sheet("Perform a reconciliation")

        # get parameters from excel
        LeftScope = "F24"
        LeftCode = "F25"
        LeftEffectiveAt = "F26"
        LeftAsAt = "F27"
        RightScope = "F28"
        RightCode = "F29"
        RightEffectiveAt = "F30"
        RightAsAt = "F31"

        left_scope = sht.range(LeftScope).value
        left_code = sht.range(LeftCode).value
        left_effective_at = get_date(LeftEffectiveAt, sht)
        left_as_at = get_date(LeftAsAt, sht)
        right_scope = sht.range(RightScope).value
        right_code = sht.range(RightCode).value
        right_effective_at = get_date(RightEffectiveAt, sht)
        right_as_at = get_date(RightAsAt, sht)

        # get validation data
        request_reconciliation = models.PortfoliosReconciliationRequest(
            left=models.PortfolioReconciliationRequest(
                portfolio_id=models.ResourceId(scope=left_scope, code=left_code),
                effective_at=left_effective_at,
                as_at=left_as_at
            ),
            right=models.PortfolioReconciliationRequest(
                portfolio_id=models.ResourceId(scope=right_scope, code=right_code),
                effective_at=right_effective_at,
                as_at=right_as_at
            ),
            instrument_property_keys=['Instrument/default/LusidInstrumentId']
        )

        reconciliation_validation_response = self.reconciliations_api.reconcile_holdings(
            request=request_reconciliation
        ).values

        # format validation data
        reconciliation_validation = [
            {
                "Instrument Uid": data.instrument_uid,
                "Left Units": data.left_units,
                "Right Units": data.right_units,
                "Units Difference": data.difference_units,
                "Left Cost Amount": data.left_cost.amount,
                "Right Cost Amount": data.right_cost.amount,
            }
            for data in reconciliation_validation_response
        ]

        # get excel data
        reconciliation_response_excel = sht.range("E39", "O49").value
        headers = reconciliation_response_excel.pop(0)

        # format excel data
        reconciliation_excel = [
            {
                "Instrument Uid": data[headers.index('Instrument Uid')],
                "Left Units": data[headers.index("Left Units")],
                "Right Units": data[headers.index("Right Units")],
                "Units Difference": data[headers.index("Units Difference")],
                "Left Cost Amount": data[headers.index("Left Cost Amount")],
                "Right Cost Amount": data[headers.index("Right Cost Amount")]
            }
            for data in reconciliation_response_excel if (data[0] != "" and data[0] is not None)
        ]
        # assert excel data exists within validation set
        [self.assertIn(sample, reconciliation_validation) for sample in reconciliation_excel]

    @unittest.skip("'upload transactions' Test not fully implemented")
    def test_upload_transactions(self):
        sht = self.get_sheet("Upload transactions")

        # get locations of parameters from excel
        scope_location = "F15"
        PortfolioCode = "F17"
        formula = "E31"

        # load parameters
        scope = sht.range(scope_location).value
        code = sht.range(PortfolioCode).value
        up_transactions_response_excel = sht.range("E19", "R29").value
        header = up_transactions_response_excel.pop(0)
        self.assertTrue(scope, msg="Error loading scope from location: " + scope_location)
        self.assertTrue(code, msg="Error loading scope from location: " + PortfolioCode)

        # get validation data

        # format validation data

        # format excel data

        # perform assertion
