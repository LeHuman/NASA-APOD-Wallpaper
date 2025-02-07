use std::ffi::OsStr;
use std::os::windows::ffi::OsStrExt;
use std::ptr;

use winapi::ctypes::wchar_t;
use winapi::shared::winerror::{FAILED, SUCCEEDED};
use winapi::um::combaseapi::{CoCreateInstance, CoInitializeEx, CoTaskMemFree, CoUninitialize};
use winapi::um::objbase::COINIT_APARTMENTTHREADED;
use winapi::um::shobjidl_core::{CLSID_DesktopWallpaper, IDesktopWallpaper};
use winapi::um::wingdi::DEVMODEW;
use winapi::Interface;

fn to_wide_str(s: &str) -> Vec<u16> {
    OsStr::new(s).encode_wide().chain(Some(0)).collect()
}

pub(crate) struct Monitor {
    id: Vec<u16>,
    width: u32,
    height: u32,
}

impl Monitor {
    pub(crate) fn size(&self) -> (u32, u32) {
        (self.width, self.height)
    }
}

pub(crate) struct Wallpaper {
    instance: *mut IDesktopWallpaper,
}

impl Wallpaper {
    pub(crate) fn new() -> Result<Self, Box<dyn std::error::Error>> {
        unsafe {
            if FAILED(CoInitializeEx(ptr::null_mut(), COINIT_APARTMENTTHREADED)) {
                return Err("Failed to initialize".into());
            }

            let mut desktop_wallpaper: *mut IDesktopWallpaper = ptr::null_mut();

            let hr = CoCreateInstance(
                &CLSID_DesktopWallpaper,
                ptr::null_mut(),
                winapi::um::combaseapi::CLSCTX_ALL,
                &IDesktopWallpaper::uuidof(),
                &mut desktop_wallpaper as *mut _ as *mut *mut winapi::ctypes::c_void,
            );

            if FAILED(hr) {
                CoUninitialize();
                return Err("Failed to create IDesktopWallpaper instance".into());
            }

            Ok(Self {
                instance: desktop_wallpaper,
            })
        }
    }

    pub(crate) fn get_monitors(&self) -> Vec<Monitor> {
        let mut monitors = Vec::new();
        unsafe {
            let mut monitor_count = 0;
            let instance = &*self.instance;
            if SUCCEEDED(instance.GetMonitorDevicePathCount(&mut monitor_count)) {
                for i in 0..monitor_count {
                    let mut monitor_id: *mut wchar_t = ptr::null_mut();
                    if SUCCEEDED(instance.GetMonitorDevicePathAt(i, &mut monitor_id)) && !monitor_id.is_null() {
                        // Convert to Rust's owned data to prevent use-after-free
                        let monitor_name_vec =
                            std::slice::from_raw_parts(monitor_id, 82).to_vec(); // FIXME: What is the actual length needed here?
                        let monitor_name = String::from_utf16_lossy(&monitor_name_vec);
                        println!("Monitor {} ID: {}", i, monitor_name);

                        // Get resolution
                        let mut dev_mode: DEVMODEW = std::mem::zeroed();
                        dev_mode.dmSize = std::mem::size_of::<DEVMODEW>() as u16;

                        let mut monitor_rect = std::mem::zeroed();
                        if SUCCEEDED(instance.GetMonitorRECT(monitor_id, &mut monitor_rect)) {
                            monitors.push(Monitor {
                                id: monitor_name_vec,
                                width: (monitor_rect.right - monitor_rect.left) as u32,
                                height: (monitor_rect.bottom - monitor_rect.top) as u32,
                            });
                        } else {
                            eprintln!("Failed to get resolution for monitor {}", i);
                        }

                        CoTaskMemFree(monitor_id as *mut _);
                    }
                }
            }
        }
        monitors
    }
    pub(crate) fn set_wallpaper(
        &self,
        path: &str,
        monitor: &Monitor,
    ) -> Result<(), Box<dyn std::error::Error>> {
        unsafe {
            let instance = &*self.instance;
            let wallpaper_path = to_wide_str(path);
            if FAILED(instance.SetWallpaper(monitor.id.as_ptr(), wallpaper_path.as_ptr())) {
                return Err("Failed to set wallpaper".into());
            }
        }
        Ok(())
    }
}

impl Drop for Wallpaper {
    fn drop(&mut self) {
        unsafe {
            CoUninitialize();
        }
    }
}
