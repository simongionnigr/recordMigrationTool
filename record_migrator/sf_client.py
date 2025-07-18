from simple_salesforce import Salesforce
import pandas as pd

class SalesforceClient:
    """
    Wrapper minimalista per autenticarsi e lanciare query SOQL.
    """
    def __init__(self, username: str, password: str, token: str, is_sandbox: bool):
        domain = "test" if is_sandbox else "login"
        self.sf = Salesforce(
            username=username,
            password=password,
            security_token=token,
            domain=domain
        )

    def run_queries(self, query_df: pd.DataFrame, col_sobj: str, col_soql: str) -> dict:
        """
        Esegue tutte le query contenute in query_df e restituisce un dict:
          { sobject_api_name: DataFrame(result) }
        """
        results = {}
        for _, row in query_df.iterrows():
            so_name = row[col_sobj]
            so_sql  = row[col_soql]
            resp    = self.sf.query_all(so_sql)
            records = resp.get("records", [])
            # Rimuovo la chiave 'attributes'
            data = [{k: v for k, v in r.items() if k != "attributes"} for r in records]
            results[so_name] = pd.DataFrame(data)
        return results
