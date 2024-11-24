use crate::discord_ipc::DiscordIpc;
use serde_json::json;
use std::{error::Error, path::PathBuf};
use windows::{
    core::PCWSTR,
    Win32::{
        Foundation::{CloseHandle, HANDLE, INVALID_HANDLE_VALUE},
        Storage::FileSystem::{
            CreateFileW,
            FlushFileBuffers,
            ReadFile,
            WriteFile,
            FILE_ATTRIBUTE_NORMAL,
            FILE_SHARE_READ,
            FILE_SHARE_WRITE,
            OPEN_EXISTING,
        },
    },
};
use windows::Win32::Foundation::{GENERIC_READ, GENERIC_WRITE};

type Result<T> = std::result::Result<T, Box<dyn Error>>;

#[allow(dead_code)]
#[derive(Debug)]
/// A wrapper struct for the functionality contained in the
/// underlying [`DiscordIpc`](trait@DiscordIpc) trait.
pub struct DiscordIpcClient {
    /// Client ID of the IPC client.
    pub client_id: String,
    connected: bool,
    pipe_handle: Option<HANDLE>,
}

impl DiscordIpcClient {
    /// Creates a new `DiscordIpcClient`.
    ///
    /// # Examples
    /// ```
    /// let ipc_client = DiscordIpcClient::new("<some client id>")?;
    /// ```
    pub fn new(client_id: &str) -> Result<Self> {
        Ok(Self {
            client_id: client_id.to_string(),
            connected: false,
            pipe_handle: None,
        })
    }

    // Add a method to check if the pipe is still valid
    unsafe fn is_pipe_valid(&self) -> bool {
        if let Some(handle) = self.pipe_handle {
            // Try to write 0 bytes to check if pipe is still connected
            let mut bytes_written = 0;
            match WriteFile(
                handle,
                Some(&[]),
                Some(&mut bytes_written),
                None,
            ) {
                Ok(_) => true,
                Err(e) => {
                    false
                },
            }
        } else {
            false
        }
    }

    // Add a method to reconnect if needed
    fn ensure_connected(&mut self) -> Result<()> {
        if !self.connected || unsafe { !self.is_pipe_valid() } {
            self.connected = false;
            // Close existing handle if any
            if let Some(handle) = self.pipe_handle {
                unsafe {
                    let _ = CloseHandle(handle);
                }
            }
            self.pipe_handle = None;
            self.connect_ipc()?;
        }
        Ok(())
    }

    unsafe fn create_pipe_handle(path: &PathBuf) -> windows::core::Result<HANDLE> {
        let wide_path = path.to_str()
            .unwrap()
            .encode_utf16()
            .chain(Some(0))
            .collect::<Vec<u16>>();

        let handle = CreateFileW(
            PCWSTR(wide_path.as_ptr()),
            GENERIC_READ.0 | GENERIC_WRITE.0,
            FILE_SHARE_READ | FILE_SHARE_WRITE,
            None,
            OPEN_EXISTING,
            FILE_ATTRIBUTE_NORMAL,
            HANDLE(std::ptr::null_mut()),
        )?;

        if handle == INVALID_HANDLE_VALUE {
            return Err(windows::core::Error::from_win32());
        }

        Ok(handle)
    }
}

impl Drop for DiscordIpcClient {
    fn drop(&mut self) {
        if let Some(handle) = self.pipe_handle {
            unsafe {
                let _ = self.close();
                let _ = CloseHandle(handle);
            }
        }
    }
}

impl DiscordIpc for DiscordIpcClient {
    fn connect_ipc(&mut self) -> Result<()> {
        for i in 0..10 {
            let path = PathBuf::from(format!(r"\\?\pipe\discord-ipc-{}", i));

            unsafe {
                match Self::create_pipe_handle(&path) {
                    Ok(handle) => {
                        self.pipe_handle = Some(handle);
                        self.connected = true;
                        return Ok(());
                    }
                    Err(e) => {
                        continue
                    },
                }
            }
        }

        Err("Could not connect to Discord IPC pipe".into())
    }

    fn write(&mut self, data: &[u8]) -> Result<()> {
        self.ensure_connected()?;


        let mut retries = 3;

        while retries > 0 {
            let handle = self.pipe_handle.ok_or("Pipe handle not initialized")?;
            let mut bytes_written = 0;

            unsafe {
                return match WriteFile(
                    handle,
                    Some(data),
                    Some(&mut bytes_written),
                    None,
                ) {
                    Ok(_) => Ok(()),
                    Err(e) => {
                        if e.code() == windows::core::HRESULT::from_win32(0x800700E8) {
                            // Pipe is being closed error
                            if retries > 1 {
                                self.connected = false;
                                self.ensure_connected()?;

                                retries -= 1;
                                continue;
                            }
                        }
                        Err(e.into())
                    }
                }
            }
        }

        Err("Failed to write to pipe after retries".into())
    }

    fn read(&mut self, buffer: &mut [u8]) -> Result<()> {
        let handle = self.pipe_handle.ok_or("Pipe handle not initialized")?;
        let mut bytes_read = 0;

        unsafe {
            ReadFile(
                handle,
                Some(buffer),
                Some(&mut bytes_read),
                None,
            ).map_err(|e| e.into())
        }
    }

    fn close(&mut self) -> Result<()> {
        let data = json!({});
        let _ = self.send(data, 2);

        if let Some(handle) = self.pipe_handle {
            unsafe {
                FlushFileBuffers(handle)?
            }
        }

        self.connected = false;
        Ok(())
    }

    fn get_client_id(&self) -> &String {
        &self.client_id
    }
}