use clap::Parser;

#[derive(Parser, Debug)]
#[command(author, version, about, long_about = None)]
pub struct Args {
    #[arg(short, long)]
    pub input_path: String,

    #[arg(short, long)]
    pub output_path: String,

    #[arg(short, long, default_value = "main")]
    pub entry_sheet: Option<String>,
}

pub fn parse_args() -> Args {
    Args::parse()
}
