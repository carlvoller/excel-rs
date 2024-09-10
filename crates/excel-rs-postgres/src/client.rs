use anyhow::Result;
use postgres::{Client, NoTls, RowIter};
use rustls::ClientConfig;
use tokio_postgres_rustls::MakeRustlsConnect;

use crate::ssl::SkipServerVerification;

pub struct PostgresClient {
    client: Client,
}

impl PostgresClient {
    pub fn new(conn_string: &str) -> PostgresClient {
        let config = ClientConfig::builder()
            .with_root_certificates(rustls::RootCertStore::empty())
            .with_no_client_auth();

        let tls = MakeRustlsConnect::new(config);

        let client = match Client::connect(conn_string, tls) {
            Ok(c) => c,
            Err(_) => {
                // Attempt SSL with Skipped Verification
                let mut config = ClientConfig::builder()
                    .with_root_certificates(rustls::RootCertStore::empty())
                    .with_no_client_auth();

                config
                    .dangerous()
                    .set_certificate_verifier(SkipServerVerification::new());

                let tls = MakeRustlsConnect::new(config);

                match Client::connect(&conn_string, tls) {
                    Ok(c) => c,
                    Err(_) => {
                        // Attempt no SSL
                        match Client::connect(&conn_string, NoTls) {
                            Ok(c) => c,
                            Err(e) => panic!("Couldn't connec to server: {e}"),
                        }
                    }
                }
            }
        };

        PostgresClient { client }
    }

    pub fn make_query(&mut self, query: &str, params: Vec<String>) -> Result<RowIter<'_>> {
        let iter: RowIter<'_> = self.client.query_raw(query, params)?;
        Ok(iter)
    }

    pub fn close(self) -> Result<()> {
        self.client.close()?;
        Ok(())
    }
}
