#!/usr/bin/env python3

import argparse
import json
import pathlib
import re

import openpyxl
import pydub

import vbs


def extract_desc(str):
    # Regex pattern
    pattern = r"(\S+)\s+(\d+)\s*-\s*(\d+)"

    # Match the pattern
    match = re.match(pattern, str)

    if not match:
        return None

    return (match.group(1), int(match.group(2)), int(match.group(3)))


def get_split_timestamps(filename):
    splits = [{"start": s[0], "end": s[1]} for s in vbs.load_and_split(filename)]

    return splits


def read_word_spec(xls_file):
    workbook = openpyxl.load_workbook(xls_file, data_only=True)

    # Select the active worksheet
    sheet = workbook.active

    entries = {}

    # First row is headers, start from second.
    for row in sheet.iter_rows(min_row=2, max_col=4, values_only=True):
        # Unpack the row values
        idx, code, ktiv_male, menukkad = row
        # Check if the row has all required values
        if all([idx, code, ktiv_male, menukkad]):
            entry = {"idx": int(idx), "code": code, "ktiv_male": ktiv_male, "menukkad": menukkad}
            entries[code] = entry
        elif any([idx, code, ktiv_male, menukkad]):
            raise ValueError(f"Incomplete data in row: {row}")

    return entries


def split(audio_files, xls_file, output_dir):
    word_spec = read_word_spec(xls_file)

    json.dump(word_spec, open(f"{output_dir}/desc.json", "w"))

    idx_to_ws = {}
    for e in word_spec:
        idx_to_ws[word_spec[e]['idx']] = word_spec[e]

    for audio_file in audio_files:
        base_filename = pathlib.Path(audio_file).stem
        desc = extract_desc(base_filename)

        if not desc:
            print("Error extracting file information. Quitting.")
            sys.exit(1)

        letter, first_idx, last_idx = desc
        print(f"Extracting audio for letter {letter}, words {first_idx}-{last_idx}...")

        splits = get_split_timestamps(audio_file)

        if len(splits) != (last_idx - first_idx + 1) + 1:
            print(f"Has {len(splits)} splits, not matching expected.")
            print(splits)

        pd_file = pydub.AudioSegment.from_file(audio_file)
        for idx, s in enumerate(splits[1:]):
            start = s["start"] - 0.3
            end = s["end"] + 0.3

            ws_entry = idx_to_ws[first_idx + idx]
            pd_file[int(start * 1000) : int(end * 1000)].export(
                f'{output_dir}/{ws_entry["code"]}_{ws_entry["ktiv_male"]}.mp3', format="mp3"
            )


if __name__ == "__main__":
    # Define an argument parser
    parser = argparse.ArgumentParser(description="Split an audio file to separate words, defined in an XLS file.")

    # Add the arguments
    parser.add_argument("--audio", action="append", required=True, help="Audio file to split.")
    parser.add_argument("--xls", type=str, required=True, help="XLS file containing word specification.")
    parser.add_argument(
        "--output-dir", type=str, required=True, help="The directory where splitted audios will be stored."
    )

    # Parse the arguments
    args = parser.parse_args()

    split(args.audio, args.xls, args.output_dir)
