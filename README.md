# Welcome to the raft (register_app_for_toast) Readme.

> [!Note]
> This crate is for Windows only.

---

This is a crate that registers your **commandline** app for toast notifications, but only on Windows.

[log](https://crates.io/crates/log) is supported to give a few insights on what's happening in the background.

**IMPORTANT !!**:
- EVERYTIME after running the lib's register function for the first time, you may need to open a new terminal for it to load the changed dependencies and registry keys.
- If opening a new terminal doesn't work, there should be a new shortcut in your start menu (with the App name you provided). Click it and it will work.

---

### Usage

---

It's designed to be used with [winrt_toast_reborn](https://crates.io/crates/winrt-toast-reborn) because the `register` function in `winrt_toast_reborn` did not work for me.

The toast functions (from [winrt_toast_reborn](https://crates.io/crates/winrt-toast-reborn)), however, worked fine.

1. Generate a GUID from any service (like [DevToys]()) or online.

2. register the app & throw some toasts
```rust
use register_app_for_toast::register;
use winrt_toast_reborn::{Action as ToastAction, Toast, ToastManager};

fn main() {
    // register the app
    let aumid = "com.your.unique.app.id"; // any unique string
    let clsid = "afe12491-6e01-42a5-bb43-d8467ff49af7"; // guid you generated
    let app_name = "Cool App Name";
    match register(
        aumid,
        clsid,
        Some(app_name), // your app name
    ) {
        Ok(_) => {
            println!("Registered {aumid} as \"{app_name}\"");
        }
        Err(e) => {
            eprintln!("Registration Error: {e}");
        }
    }
    // then use a toast with some actions
    let manager = ToastManager::new(app_aumid);

    let mut toast = Toast::new();
    toast
        .text1("Toast Title!")
        .action(ToastAction::new("Close", "close", "close"));

    manager.show(&toast);
}
```

3. In your custom **uninstaller**, you can use the `deregister` function:
```rust
// uninstaller.rs
use register_app_for_toast::deregister;

fn main() {
    // register the app
    let aumid = "com.your.unique.app.id"; // your app unique id
    let clsid = "afe12491-6e01-42a5-bb43-d8467ff49af7"; // your clsid
    let app_name = "Cool App Name";
    match deregister(
        aumid,
        clsid,
        Some(app_name), // your app name
    ) {
        Ok(_) => {
            println!("Deregistered {aumid} from system");
        }
        Err(e) => {
            eprintln!("Deregistration Error: {e}");
        }
    }
}
```
