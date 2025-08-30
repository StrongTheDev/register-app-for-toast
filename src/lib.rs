#[cfg(windows)]
pub use windows_module::*;

#[cfg(windows)]
mod windows_module {
    use anyhow::{anyhow, Result};
    use log::{debug, error, info};
    use std::{env, path::Path};
    use windows::{
        core::*,
        Win32::{
            Foundation::*,
            System::Com::{StructuredStorage::PROPVARIANT, *},
            UI::Shell::*,
        },
    };
    use winreg::{enums::HKEY_CURRENT_USER, RegKey};

    fn handle_shortcuts(
        app_aumid: &str,
        clsid_str: &str,
        app_name: &str,
        is_registering: bool,
    ) -> Result<String> {
        // Path to the Start Menu shortcut directory
        let shortcut_path = format!(
            r"{}\Microsoft\Windows\Start Menu\Programs\{app_name}.lnk",
            env!("APPDATA")
        );

        if !is_registering {
            let path = Path::new(&shortcut_path);
            if !path.exists() {
                error!("The shortcut ({shortcut_path}) has not been found!");
                return Err(anyhow!(
                    "The shortcut ({shortcut_path}) has not been found!"
                ));
            }
            debug!("Deleting shortcut ...");
            std::fs::remove_file(path)?;
            info!("Completed ...!");
            return Ok("".into());
        }

        debug!("Creating shortcut with name ({app_name}) AUMID: {app_aumid}");
        let app_path = env::current_exe()?.to_string_lossy().into_owned();

        unsafe {
            let shell_link: IShellLinkW = CoCreateInstance(&ShellLink, None, CLSCTX_INPROC_SERVER)?;
            let h = HSTRING::from(&app_path);
            let h = h.as_ptr();
            let pszfile = PCWSTR::from_raw(h);
            // info!("PCW: {}", pszfile.to_string()?)

            const GUID_FOR_TOASTS: GUID = GUID::from_u128(0x9F4C2855_9F79_4B39_A8D0_E1D42DE1D5F3);
            shell_link.SetPath(pszfile)?;
            let property_store: PropertiesSystem::IPropertyStore = shell_link.cast()?;
            // aumid property key
            let property_key_aumid = PROPERTYKEY {
                fmtid: GUID_FOR_TOASTS,
                pid: 5,
            };
            let var_aumid = PROPVARIANT::from(app_aumid);
            property_store.SetValue(
                &property_key_aumid as *const PROPERTYKEY,
                &var_aumid as *const PROPVARIANT,
            )?;

            // toast activator
            let property_key_activator = PROPERTYKEY {
                fmtid: GUID_FOR_TOASTS,
                pid: 26,
            };
            let var_activator: &str = &format!("{{{}}}", clsid_str);
            let var_activator = PROPVARIANT::from(var_activator);
            property_store.SetValue(
                &property_key_activator as *const PROPERTYKEY,
                &var_activator as *const PROPVARIANT,
            )?;
            property_store.Commit()?;

            let persist_file: IPersistFile = shell_link.cast()?;
            let h = HSTRING::from(&shortcut_path);
            let h = h.as_ptr();
            let shortcut_path = PCWSTR::from_raw(h);
            persist_file.Save(shortcut_path, true)?;
        }
        info!("Shortcut created at: {}", shortcut_path);

        Ok(shortcut_path)
    }

    fn handle_com_server(clsid_str: &str, is_registering: bool) -> Result<()> {
        let hkcu = RegKey::predef(HKEY_CURRENT_USER);
        let clsid_path = format!(r"Software\Classes\CLSID\{{{}}}", clsid_str);

        if !is_registering {
            debug!("Deleting COM server with CLSID: {}", clsid_str);
            if let Err(e) = hkcu.delete_subkey_all(&clsid_path) {
                error!("Failed to delete registry key ({clsid_path}) in HKCU\nError: {e}");
                return Err(anyhow!(
                    "Failed to delete registry key ({clsid_path}) in HKCU\nError: {e}"
                ));
            }
            info!("COM server deleted successfully");
            return Ok(());
        }

        let appid_path = format!(r"Software\Classes\AppID\{{{}}}", clsid_str);
        debug!("Registering COM server with CLSID: {}", clsid_str);
        // Register CLSID under HKCU\Software\Classes\CLSID
        let clsid_key = hkcu.create_subkey(&clsid_path)?.0;
        // Set AppID
        clsid_key.set_value("AppID", &format!("{}", clsid_str))?;
        // Creae LocalServer32 subkey
        let (local_server_key, _) = clsid_key.create_subkey("LocalServer32")?;
        let exe_path = std::env::current_exe()?.to_string_lossy().into_owned();
        local_server_key.set_value("", &format!("\"{}\" -ToastActivated", exe_path))?;
        // Register AppID
        let (appid_key, _) = hkcu.create_subkey(&appid_path)?;
        appid_key.set_value("DllSurrogate", &"")?;
        info!("COM server registered successfully");

        Ok(())
    }

    fn handle_notification_support(
        app_aumid: &str,
        clsid_str: &str,
        app_name: &str,
        is_registering: bool,
    ) -> Result<String> {
        // Initialize COM
        unsafe {
            CoInitializeEx(None, COINIT_APARTMENTTHREADED).unwrap();
        }
        handle_com_server(clsid_str, is_registering)?;

        let shortcut = handle_shortcuts(app_aumid, clsid_str, app_name, is_registering)?;

        unsafe {
            CoUninitialize();
        }
        Ok(shortcut)
    }

    /// Registers the application to be able to send toast notifications.
    ///
    /// # Arguments
    ///
    /// * `aumid` - The Application User Model ID for the application. MUST be unique e.g. `com.developer.app.submodule` or `MyCompany.MyApp.Module`
    /// * `clsid` - The CLSID of the COM server that will handle the toast notification activation. A clsid is basically a GUID string. Many generators exist online and offline e.g. [**DevToys** (Microsoft Store)](https://apps.microsoft.com/detail/9PGCV4V3BK4W?hl=en&gl=UG&ocid=pdpshare) (Under Generators)
    /// * `app_name` - The name of the application to be displayed in the shortcut. If not provided, defaults to use the aumid as the app name.
    ///
    /// @Returns the shortcut path
    ///
    /// Example:
    /// ```rust
    /// let _shortcut_path: String = register(
    ///     "cool.unique.app.identifier", // unique e.g. reverse domain name
    ///     "afe12481-6e51-42a1-bc43-d8487fe49af8", // your unique app GUID/CLSID
    ///     Some("Cool App Name"), // The name that will be displayed on your app shortcut/notification
    /// ).unwrap();
    /// ```
    pub fn register(aumid: &str, clsid: &str, app_name: Option<&str>) -> Result<String> {
        let app_name = if app_name.is_some() {
            app_name.unwrap()
        } else {
            aumid
        };
        let shortcut_path = handle_notification_support(aumid, clsid, app_name, true)?;
        Ok(shortcut_path)
    }

    /// Removes application ability to send toast notifications.
    ///
    /// # Arguments
    /// **NB:** `aumid`, `clsid` and `app_name` **MUST** be the same as the ones used in the [register] command; in order to get the correct results.
    ///
    /// * `aumid` - The Application User Model ID for the application e.g. `com.developer.app.submodule` or `MyCompany.MyApp.Module`
    /// * `clsid` - The CLSID of the COM server that will handle the toast notification activation. A clsid is basically a GUID string.
    /// * `app_name` - The name of the application as displayed in the shortcut. If not provided, defaults to use the aumid as the app name.
    ///
    /// Example:
    /// ```rust
    /// deregister(
    ///     "cool.unique.app.identifier", // your app's unique identifier name
    ///     "afe12481-6e51-42a1-bc43-d8487fe49af8", // your unique app GUID/CLSID
    ///     Some("Cool App Name"), // The name you used for your app shortcut
    /// ).unwrap();
    /// ```
    pub fn deregister(aumid: &str, clsid: &str, app_name: Option<&str>) -> Result<()> {
        let app_name = if app_name.is_some() {
            app_name.unwrap()
        } else {
            aumid
        };
        handle_notification_support(aumid, clsid, app_name, false)?;
        Ok(())
    }
}
