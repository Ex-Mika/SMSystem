USE `school_management_system`;

CREATE TABLE IF NOT EXISTS `student_subject_grades` (
  `grade_id` INT NOT NULL AUTO_INCREMENT,
  `teacher_id` INT NOT NULL,
  `student_id` INT NOT NULL,
  `subject_id` INT NOT NULL,
  `section_name` VARCHAR(60) NULL,
  `quiz_score` DECIMAL(5,2) NOT NULL DEFAULT 0.00,
  `project_score` DECIMAL(5,2) NOT NULL DEFAULT 0.00,
  `midterm_score` DECIMAL(5,2) NOT NULL DEFAULT 0.00,
  `final_exam_score` DECIMAL(5,2) NOT NULL DEFAULT 0.00,
  `final_grade` DECIMAL(5,2) NOT NULL DEFAULT 0.00,
  `remarks` VARCHAR(30) NULL,
  `created_at` DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
  `updated_at` DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  PRIMARY KEY (`grade_id`),
  UNIQUE KEY `uq_student_subject_grades_teacher_student_subject` (`teacher_id`, `student_id`, `subject_id`),
  KEY `ix_student_subject_grades_student` (`student_id`),
  KEY `ix_student_subject_grades_subject` (`subject_id`),
  CONSTRAINT `fk_student_subject_grades_teacher`
    FOREIGN KEY (`teacher_id`) REFERENCES `teachers` (`teacher_id`)
    ON DELETE CASCADE,
  CONSTRAINT `fk_student_subject_grades_student`
    FOREIGN KEY (`student_id`) REFERENCES `students` (`student_id`)
    ON DELETE CASCADE,
  CONSTRAINT `fk_student_subject_grades_subject`
    FOREIGN KEY (`subject_id`) REFERENCES `subjects` (`subject_id`)
    ON DELETE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
