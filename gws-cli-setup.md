# GWS CLI Setup

Installed via Cargo from source:

```bash
cargo install --git https://github.com/googleworkspace/cli --locked
```

Version: 0.22.5  
Binary: `~/.cargo/bin/gws`

## Next Steps

Run authentication setup:

```bash
gws auth setup
gws auth login
```

Requires a Google Cloud project and active Google Workspace account.

## Example Usage

```bash
gws drive files list --params '{"pageSize": 5}'
gws gmail users messages list --params '{"userId": "me"}'
gws sheets spreadsheets get --params '{"spreadsheetId": "..."}'
```
