use std::path::PathBuf;

use dcc_mcp_office_mcp_server::OfficeMcpServer;
use rmcp::{transport::stdio, ServiceExt};

const USAGE: &str = "usage: dcc-mcp-office-mcp-server --app=<powerpoint|word|excel> [--workspace-root=<path>] [--host=<dcc-office-host.exe>]";

struct Arguments {
    app: String,
    workspace_root: PathBuf,
    host: Option<PathBuf>,
}

#[tokio::main]
async fn main() {
    match run().await {
        Ok(()) => {}
        Err(error) => {
            eprintln!("dcc-mcp-office-mcp-server: {error}");
            std::process::exit(1);
        }
    }
}

async fn run() -> Result<(), Box<dyn std::error::Error>> {
    let Some(arguments) = parse_arguments(std::env::args().skip(1))? else {
        return Ok(());
    };
    let server = OfficeMcpServer::start(
        &arguments.app,
        &arguments.workspace_root,
        arguments.host.as_deref(),
    )
    .await?;
    let service = server.serve(stdio()).await?;
    service.waiting().await?;
    Ok(())
}

fn parse_arguments(
    arguments: impl IntoIterator<Item = String>,
) -> Result<Option<Arguments>, String> {
    let mut app = None;
    let mut workspace_root = None;
    let mut host = None;
    for argument in arguments {
        if matches!(argument.as_str(), "--help" | "-h") {
            println!("{USAGE}");
            return Ok(None);
        }
        if argument == "--version" {
            println!("dcc-mcp-office-mcp-server {}", env!("CARGO_PKG_VERSION"));
            return Ok(None);
        }
        if let Some(value) = argument.strip_prefix("--app=") {
            app = Some(value.to_string());
        } else if let Some(value) = argument.strip_prefix("--workspace-root=") {
            workspace_root = Some(PathBuf::from(value));
        } else if let Some(value) = argument.strip_prefix("--host=") {
            host = Some(PathBuf::from(value));
        } else {
            return Err(format!("unknown argument '{argument}'. {USAGE}"));
        }
    }
    let app = app.ok_or_else(|| format!("--app is required. {USAGE}"))?;
    if !matches!(app.as_str(), "powerpoint" | "word" | "excel") {
        return Err(format!("unsupported --app '{app}'. {USAGE}"));
    }
    let workspace_root = match workspace_root {
        Some(path) => path,
        None => std::env::current_dir().map_err(|error| error.to_string())?,
    };
    Ok(Some(Arguments {
        app,
        workspace_root,
        host,
    }))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn cli_requires_an_explicit_supported_app() {
        assert!(parse_arguments(Vec::<String>::new()).is_err());
        assert!(parse_arguments(["--app=powerpoint".into()])
            .unwrap()
            .is_some());
        assert!(parse_arguments(["--app=outlook".into()]).is_err());
    }
}
