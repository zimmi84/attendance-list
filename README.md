# attendance-list
with this program you can create an excel for a training attendance list

## Usage

It is necessary to have two files on the same path prepared.
- teamDa.xlsx
- teamDb.xlsx

Each file must have a row "surname" and "lastname" otherwise the python script is not able to import the attendance/players to your attendance list.

```bash
python create-attendance-list.py <start_date> <end_date> [extra_day_start extra_day_end]
```

start_date - mandatory
end_date - mandatory

extra_day_start - optinal
extra_day end - optional