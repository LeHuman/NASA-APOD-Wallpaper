use ab_glyph::{FontRef, PxScale};
use directories::UserDirs;
use image::{imageops::FilterType, DynamicImage, GenericImageView};
use imageproc::drawing::draw_text_mut;
use reqwest::blocking::get;
use reqwest::Url;
use serde::Deserialize;
use std::cmp::max;
use std::cmp::min;
use std::env;
use std::fs;
use std::fs::File;
use std::io::copy;
use std::net::TcpStream;
use std::path::Path;
use std::path::PathBuf;
use std::time::Duration;
use std::time::Instant;

const WATERMARK: &[u8] = include_bytes!("watermark.png");

#[allow(dead_code)]
const FONT_LICENSE: &[u8] = include_bytes!("font/LICENSE");
const TITLE_FONT: &[u8] = include_bytes!("font/IBMPlexSans-Bold.ttf");
const FONT: &[u8] = include_bytes!("font/IBMPlexSans-Text.ttf");

mod wallpaper;
use wallpaper::Wallpaper;

#[derive(Debug, Deserialize, PartialEq, Eq)]
pub struct ApodResponse {
    pub title: String,
    pub explanation: String,
    pub copyright: Option<String>,
    pub url: String,
    #[serde(rename = "hdurl")]
    pub hd_url: Option<String>,
    pub media_type: String,
}

pub fn get_nasa_apod_folder() -> Option<PathBuf> {
    if let Some(user_dirs) = UserDirs::new() {
        if let Some(pictures_dir) = user_dirs.picture_dir() {
            let nasa_apod_path = pictures_dir.join("NASA_APOD");

            if !nasa_apod_path.exists() {
                if let Err(e) = fs::create_dir_all(&nasa_apod_path) {
                    eprintln!("Failed to create directory: {}", e);
                    return None;
                }
            }

            return Some(nasa_apod_path);
        }
    }
    None
}

fn check_connection(timeout: Duration) -> Result<(), String> {
    let deadline = Instant::now() + timeout;

    loop {
        match TcpStream::connect("api.nasa.gov:80") {
            Ok(_) => {
                println!("Internet connection detected");
                return Ok(());
            }
            Err(_) => {
                if Instant::now() >= deadline {
                    return Err("Timeout reached, no connection".into());
                }
                std::thread::sleep(Duration::from_secs(1));
            }
        }
    }
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let args: Vec<String> = env::args().collect();

    if args.len() < 2 {
        println!("Usage: {} [NASA_API_KEY]", args[0]);
        return Ok(());
    }

    let save_directory = get_nasa_apod_folder().ok_or("Failed to get APOD pictures folder")?;

    let input = &args[1];

    let timeout = Duration::from_secs(30);
    println!("Waiting for internet connection");
    check_connection(timeout)?;

    match fetch_apod(format!(
        "https://api.nasa.gov/planetary/apod?api_key={}&hd=true",
        input
    )) {
        Ok(apod) => {
            if let Err(e) = process_image(&save_directory, &apod) {
                eprintln!("Error processing image: {}", e);
            }
        }
        Err(e) => eprintln!("Failed to fetch APOD: {}", e),
    }
    Ok(())
}

#[test]
fn test_local_apod() -> Result<(), Box<dyn std::error::Error>> {
    let save_dir = format!("{}/NASA_APOD", env::temp_dir().to_str().unwrap());
    fs::create_dir_all(&save_dir)?;
    let mut img = image::open("NASA_APOD/test_apod.jpg")?;
    let wallpaper = Wallpaper::new()?;
    let monitors = wallpaper.get_monitors();

    if monitors.is_empty() {
        return Err("No monitors detected.".into());
    }

    let mut already_watermarked = false;

    for (index, monitor) in monitors.iter().enumerate() {
        let (width, height) = monitor.size();
        let resized_img = resize_and_sharpen(&mut img, width, height)?;
        let scale = ((1920.0 / f64::from(width)) + (1080.0 / f64::from(height))) / 2.0;
        let annotated_img = add_text_overlay(
            scale as f32,
            resized_img,
            "A very epic title",
            "This is so epic",
        )?;
        let mut final_img = annotated_img;
        if !already_watermarked {
            final_img = add_watermark(final_img.clone(), width, height)?;
        }
        already_watermarked = true;

        let output_filename = format!("{}/wallpaper_{}.png", save_dir, index);
        final_img.save_with_format(&output_filename, image::ImageFormat::Png)?;
        println!("Saved: {}", output_filename);

        wallpaper.set_wallpaper(&output_filename, monitor)?;
    }

    Ok(())
}

fn fetch_apod(apod_url: String) -> Result<ApodResponse, Box<dyn std::error::Error>> {
    let response = get(apod_url)?;
    let response = response.json::<ApodResponse>()?;

    if response.media_type != "image" {
        return Err("NASA APOD is not an image today.".into());
    }

    Ok(response)
}

fn process_image(save_dir: &Path, apod: &ApodResponse) -> Result<(), Box<dyn std::error::Error>> {
    let wallpaper = Wallpaper::new().unwrap();
    let monitors = wallpaper.get_monitors();

    if monitors.is_empty() {
        return Err("No monitors detected.".into());
    }

    let filepath = download_image(&apod.hd_url.clone().unwrap_or(apod.url.clone()), save_dir)?;
    let mut img = image::open(&filepath)?;

    let watermarked_monitor = rand::random_range(0..monitors.len());

    for (index, monitor) in monitors.iter().enumerate() {
        let (width, height) = monitor.size();
        let resized_img = resize_and_sharpen(&mut img, width, height)?;
        let scale = ((f64::from(width) / 1920.0) + (f64::from(height) / 1080.0)) / 2.0;
        let annotated_img =
            add_text_overlay(scale as f32, resized_img, &apod.title, &apod.explanation)?;
        let mut final_img = annotated_img;
        if index == watermarked_monitor {
            println!("Watermarked: {}", watermarked_monitor);
            final_img = add_watermark(final_img, width, height)?;
        }

        let output_filepath = save_dir.join(format!("Monitor{}.png", index));
        let output_filename = output_filepath.to_str().unwrap_or_default();

        final_img.save_with_format(&output_filepath, image::ImageFormat::Png)?;
        println!("Saved Image: {}", output_filename);

        wallpaper.set_wallpaper(output_filename, monitor)?;
    }

    Ok(())
}

fn download_image(url: &str, download_dir: &Path) -> Result<PathBuf, Box<dyn std::error::Error>> {
    let url = Url::parse(url)?;
    let filename = url
        .path_segments()
        .ok_or("Failed to split URL")?
        .last()
        .ok_or("Failed to get filename from URL")?;
    let filepath = download_dir.join(filename);

    if filepath.exists() {
        println!("Image already downloaded: {}", url);
        return Ok(filepath);
    }

    let mut response = get(url.clone())?;
    let mut file = File::create(&filepath)?;
    copy(&mut response, &mut file)?;
    println!("Downloaded Image: {}", url);
    Ok(filepath)
}

fn resize_and_sharpen(
    img: &mut DynamicImage,
    width: u32,
    height: u32,
) -> Result<DynamicImage, Box<dyn std::error::Error>> {
    let monitor_aspect_ratio = f64::from(width) / f64::from(height);
    let image_aspect_ratio = f64::from(img.width()) / f64::from(img.height());
    let crop_height;
    let crop_width;

    if image_aspect_ratio > monitor_aspect_ratio {
        // Crop by width
        crop_height = f64::from(img.height());
        crop_width = crop_height * monitor_aspect_ratio;
    } else {
        // Crop by height
        crop_width = f64::from(img.width());
        crop_height = crop_width / monitor_aspect_ratio;
    }

    let cropped = img.crop_imm(0, 0, crop_width as u32, crop_height as u32);
    let resized = cropped.resize_exact(width, height, FilterType::Lanczos3);
    // TODO: Sharpen here
    Ok(resized)
}

fn compute_image_brightness(img: &DynamicImage) -> u32 {
    let grayscale = img.to_luma8();
    let total_brightness: u32 = grayscale.pixels().map(|p| p.0[0] as u32).sum();
    let avg_brightness = total_brightness / (grayscale.width() * grayscale.height());
    avg_brightness
}

fn exponential_shrink(value: i64, factor: f64) -> i64 {
    if value == 0 {
        return 0;
    }

    let sign = value.signum();
    let magnitude = (value.abs() as f64).ln() / factor;

    (sign as f64 * magnitude.exp()).round() as i64
}

fn add_text_overlay(
    scale: f32,
    mut img: DynamicImage,
    title: &str,
    desc: &str,
) -> Result<DynamicImage, Box<dyn std::error::Error>> {
    let title_font = FontRef::try_from_slice(TITLE_FONT)?;
    let title_scale = PxScale::from(64.0 * scale);
    let font = FontRef::try_from_slice(FONT)?;
    let font_scale = PxScale::from(24.0 * scale);

    let (twidth, theight) = imageproc::drawing::text_size(title_scale, &title_font, title);
    let img_sample = img.crop_imm(
        0,
        0,
        (f64::from(twidth) * 1.2) as u32,
        (f64::from(theight) * 3.0) as u32,
    );

    const U8_MAX: u32 = u8::MAX as u32;
    let sample_brightness = compute_image_brightness(&img_sample);
    let brightness = if sample_brightness < U8_MAX { 255 } else { 0 };
    let text_color = image::Rgba([brightness, brightness, brightness, 255]);

    draw_text_mut(
        &mut img,
        text_color,
        (42.0 * scale) as i32,
        (32.0 * scale) as i32,
        title_scale,
        &title_font,
        title,
    );
    let mut lines: Vec<String> = Vec::new();
    let tokens = desc.split_whitespace();
    let max_length = 100;
    let mut max_sentence = 2;
    let mut curr_str = String::default();

    for token in tokens {
        if token.contains('.') {
            max_sentence -= 1;
        }
        if (token.len() + curr_str.len()) >= max_length {
            lines.push(curr_str.clone());
            curr_str.clear();
        }
        curr_str += " ";
        curr_str += token;
        if max_sentence == 0 {
            break;
        }
    }
    if !curr_str.is_empty() {
        lines.push(curr_str);
    }

    for (i, line) in lines.iter().enumerate() {
        draw_text_mut(
            &mut img,
            text_color,
            (40.0 * scale) as i32,
            (100.0 * scale + (28.0 * scale * (f64::from(i as u32)) as f32)) as i32,
            font_scale,
            &font,
            line,
        );
    }

    Ok(img)
}

fn add_watermark(
    mut img: DynamicImage,
    mw: u32,
    _mh: u32,
) -> Result<DynamicImage, Box<dyn std::error::Error>> {
    let image = image::load_from_memory(WATERMARK)?;

    let size = f64::from(mw) / 2.8;
    let ratio = f64::from(image.height()) / f64::from(image.width());

    let image = image.resize(size as u32, (size * ratio) as u32, FilterType::Lanczos3);
    let img_sample = img.crop_imm(
        min(img.width() - image.width(), img.width()),
        min(img.height() - image.height(), img.height()),
        min(image.width(), img.width()),
        min(image.height(), img.height()),
    );
    let image_b = compute_image_brightness(&image);
    let img_b = compute_image_brightness(&img_sample);
    let dv_b = exponential_shrink((img_b as i64) - (image_b as i64), 1.35);
    let image = image.brighten(dv_b as i32);
    let (w, h) = img.dimensions();
    let (wm_w, wm_h) = image.dimensions();

    image::imageops::overlay(&mut img, &image, (w - wm_w).into(), (h - wm_h).into());
    Ok(img)
}
