use ab_glyph::{FontRef, PxScale};
use image::{imageops::FilterType, DynamicImage, GenericImageView};
use imageproc::drawing::draw_text_mut;
use reqwest::blocking::get;
use serde::Deserialize;
use std::env;
use std::fs;
use std::fs::File;
use std::io::copy;
use std::path::Path;

const WATERMARK_PATH: &str = "dog.png"; // Path to watermark

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

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // run_test()?;
    let args: Vec<String> = env::args().collect();

    if args.len() < 2 {
        println!("Usage: {} [NASA_API_KEY]", args[0]);
        return Ok(());
    }

    let input = &args[1];

    match fetch_apod(format!(
        "https://api.nasa.gov/planetary/apod?api_key={}&hd=true",
        input
    )) {
        Ok(apod) => {
            if let Err(e) = process_image(&apod) {
                eprintln!("Error processing image: {}", e);
            }
        }
        Err(e) => eprintln!("Failed to fetch APOD: {}", e),
    }
    Ok(())
}

#[test]
fn test_local_apod() -> Result<(), Box<dyn std::error::Error>> {
    let save_dir = format!("{}/NASA_APOD", env::current_dir()?.to_str().unwrap());
    fs::create_dir_all(&save_dir)?;
    let mut img = image::open("NASA_APOD/test_apod.jpg")?;
    let wallpaper = Wallpaper::new()?;
    let monitors = wallpaper.get_monitors();

    if monitors.is_empty() {
        return Err("No monitors detected.".into());
    }

    let mut already_watermarked = false;

    for (index, monitor) in monitors.iter().enumerate() {
        let (width, height) = monitor.size().into();
        let resized_img = resize_and_sharpen(&mut img, width, height)?;
        let annotated_img = add_text_overlay(resized_img, "A very epic title", "This is so epic")?;
        let mut final_img = annotated_img;
        if !already_watermarked {
            final_img = add_watermark(final_img.clone(), width, height)?;
        }
        already_watermarked = true;

        let output_filename = format!("{}/wallpaper_{}.png", save_dir, index);
        final_img.save_with_format(&output_filename, image::ImageFormat::Png)?;
        println!("Saved: {}", output_filename);

        wallpaper.set_wallpaper(&output_filename, &monitor)?;
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

fn process_image(apod: &ApodResponse) -> Result<(), Box<dyn std::error::Error>> {
    let wallpaper = Wallpaper::new().unwrap();
    let monitors = wallpaper.get_monitors();

    if monitors.is_empty() {
        return Err("No monitors detected.".into());
    }

    // let save_dir = format!("{}\\NASA_APOD", env::var("APPDATA")?);
    let save_dir = format!("{}\\NASA_APOD", env::current_dir()?.to_str().unwrap());
    fs::create_dir_all(&save_dir)?;

    let filename = format!("{}/apod.jpg", save_dir);
    if !Path::new(&filename).exists() {
        download_image(&apod.url, &filename)?;
    }

    let mut img = image::open(&filename)?;

    let watermarked_monitor = rand::random_range(0..monitors.len());

    for (index, monitor) in monitors.iter().enumerate() {
        let (width, height) = monitor.size().into();
        let resized_img = resize_and_sharpen(&mut img, width, height)?;
        let annotated_img = add_text_overlay(resized_img, &apod.title, &apod.explanation)?;
        let mut final_img = annotated_img;
        if index == watermarked_monitor {
            println!("Watermarked: {}", watermarked_monitor);
            final_img = add_watermark(final_img, width, height)?;
        }

        let output_filename = format!("{}/wallpaper_{}.png", save_dir, index);
        final_img.save_with_format(&output_filename, image::ImageFormat::Png)?;
        println!("Saved: {}", output_filename);

        wallpaper.set_wallpaper(&output_filename, &monitor)?;
    }

    Ok(())
}

fn download_image(url: &str, filename: &str) -> Result<(), Box<dyn std::error::Error>> {
    let mut response = get(url)?;
    let mut file = File::create(filename)?;
    copy(&mut response, &mut file)?;
    Ok(())
}

fn resize_and_sharpen(
    img: &mut DynamicImage,
    width: u32,
    height: u32,
) -> Result<DynamicImage, Box<dyn std::error::Error>> {
    let monitorAspectRatio = f64::from(width) / f64::from(height);
    let imageAspectRatio = f64::from(img.width()) / f64::from(img.height());
    let mut crop_height = 0.0;
    let mut crop_width = 0.0;

    if imageAspectRatio > monitorAspectRatio {
        // Crop by width
        crop_height = f64::from(img.height());
        crop_width = f64::from(crop_height) * monitorAspectRatio;
    } else {
        // Crop by height
        crop_width = f64::from(img.width());
        crop_height = f64::from(crop_width) / monitorAspectRatio;
    }

    let cropped = img.crop(0, 0, crop_width as u32, crop_height as u32);
    let resized = cropped.resize_exact(width, height, FilterType::Lanczos3);
    // TODO: Sharpen here
    Ok(resized)
}

fn compute_image_brightness(img: &DynamicImage) -> u8 {
    let grayscale = img.to_luma8();
    let total_brightness: u32 = grayscale.pixels().map(|p| p.0[0] as u32).sum();
    let avg_brightness = total_brightness / (grayscale.width() * grayscale.height()) as u32;
    avg_brightness as u8
}

fn brightness_transform(x: u8) -> u8 {
    let normalized = x as f64 / 255.0;
    let exponent = 1.5;
    let transformed = if normalized < 0.5 {
        normalized.powf(exponent)
    } else {
        1.0 - (1.0 - normalized).powf(exponent)
    };
    (transformed * 255.0).round() as u8
}

#[test]
fn test_brightness_transformation() {
    for i in (0..=255).step_by(25) {
        println!("{} -> {}", i, brightness_transform(i));
    }
}

fn add_text_overlay(
    mut img: DynamicImage,
    title: &str,
    desc: &str,
) -> Result<DynamicImage, Box<dyn std::error::Error>> {
    let font_data = include_bytes!("HelveticaNeueBold.ttf");
    let font = FontRef::try_from_slice(font_data)?;
    let scale = PxScale::from(100.0);

    let (twidth, theight) = imageproc::drawing::text_size(scale, &font, title);
    let title_iso = img.crop(
        0,
        0,
        (f64::from(twidth) * 1.2) as u32,
        (f64::from(theight) * 3.0) as u32,
    );
    // title_iso.save_with_format("./title_sample.png", image::ImageFormat::Png)?;

    let brightness = brightness_transform(
        (compute_image_brightness(&title_iso) as u16 + u16::from(u8::MAX / 2)) as u8,
    );
    let text_color = image::Rgba([brightness, brightness, brightness, 255]);

    draw_text_mut(&mut img, text_color, 48, 32, scale, &font, title);
    let mut lines: Vec<String> = Vec::new();
    let tokens = desc.split_whitespace();
    let max_length = 100;
    let mut max_sentence = 2;
    let mut curr_str = String::default();

    for token in tokens {
        if token.contains('.') {
            max_sentence -= 1;
        }
        if ((token.len() + curr_str.len()) >= max_length) {
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
            32,
            148 + (48 * (i as i32)),
            PxScale::from(40.0),
            &font,
            line,
        );
    }

    Ok(img)
}

fn add_watermark(
    mut img: DynamicImage,
    mw: u32,
    mh: u32,
) -> Result<DynamicImage, Box<dyn std::error::Error>> {
    let image = image::open(WATERMARK_PATH)?;

    let size = f64::from(mw) / 2.8;
    let ratio = f64::from(image.height()) / f64::from(image.width());

    let watermark = image.resize(size as u32, (size * ratio) as u32, FilterType::Lanczos3);
    let (w, h) = img.dimensions();
    let (wm_w, wm_h) = watermark.dimensions();

    image::imageops::overlay(&mut img, &watermark, (w - wm_w).into(), (h - wm_h).into());
    Ok(img)
}
