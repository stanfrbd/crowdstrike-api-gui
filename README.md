# GUI for Crowdstrike API sample app

Inspired from this repo from Microsoft: https://github.com/microsoft/mde-api-gui/

> [!IMPORTANT]
> This project has nothing to do with Microsoft nor Crowdstrike.

<img width="947" height="827" alt="image" src="https://github.com/user-attachments/assets/53e8c60f-fa5c-4349-9676-c7180be3b9fc" />

## Pros

- No installation of FalconSDK needed
- Quick to execute and simple GUI
- Very useful in case of critical incident

## Cons

- Will be more difficult to keep up to date

## Scopes for Crowdstrike API Client

- Hosts (Read / Write)

## Implementation

- Connect to Crowdstrike API client and get token with good CID scope: ✅
- Add / Remove Tag: implemented and tested ✅
- Bulk Network Containment: implemented and tested but needs additional tests, especially on large CSV files (100+ hostnames) ⚠️
- Bulk Lift Containment: implemented and tested but needs additional tests, especially on large CSV files (100+ hostnames) ⚠️
- Logging: to improve ⚠️
- Creds storage: SecureString (DPAPI) so shouldn't be stored at all because of the scope of these API credentials.

## Docs

- https://falcon.eu-1.crowdstrike.com/documentation/page/c0b16f1b/host-and-host-group-management-apis
- https://www.falconpy.io/Service-Collections/Hosts.html

