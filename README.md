#### Automate the Boring Stuff with Python

This repository contains the code for automatically prepare seatplans for exams. It will read the student list and room names from csv files and generate a seatplan for each room along with a signatur sheet for the room supervisor. The result will be saved as docx files.

#### Data requirements

The student list must be a csv file with the following columns:

- `ID`: Student ID
- `Name`: Student name
- `Section`: Section number

The room list must be a csv file with the following columns:

- `Rooms`: Room name

The csv files must be saved in the `data` folder.

In addition, the configuration file `config.yaml` must be updated with the correct file names and other parameters.

- course_code: Short course code
- exam_type: Exam type (Midterm/Final)
- semester: Semester (Fall/Spring/Summer)
- year: Year

For example, if the course code is `CSE110`, exam type is `Midterm`, semester is `Fall` and year is `2019`, then the configuration file should look like this:

```yaml
rooms_file_path: data/rooms.csv
students_file_path: data/students.csv
course_code: CSE110
exam_type: Midterm
semester: Fall
year: 2019
```

#### Usage

Create a folder named `data` in the root directory and save the csv files in it as described above. Update the configuration file with the correct file names and other parameters. And finally, run the code.

To run the code, simply execute the following command:

```bash
python main.py --config config.yaml
```

If everything goes well, the seatplans will be saved in the `output` folder.

#### Dependencies

- Python 3.6+
- python-docx
- pandas
- pyyaml

#### Author

- [Mir Sazzat Hossain](https://mirsaazzathossain.me)

#### License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
