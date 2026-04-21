USE `school_management_system`;

CREATE TABLE IF NOT EXISTS `student_subject_enrollments` (
  `enrollment_id` INT NOT NULL AUTO_INCREMENT,
  `student_id` INT NOT NULL,
  `subject_id` INT NOT NULL,
  `created_at` DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
  `updated_at` DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  PRIMARY KEY (`enrollment_id`),
  UNIQUE KEY `uq_student_subject_enrollments_student_subject` (`student_id`, `subject_id`),
  KEY `ix_student_subject_enrollments_subject` (`subject_id`),
  CONSTRAINT `fk_student_subject_enrollments_student`
    FOREIGN KEY (`student_id`) REFERENCES `students` (`student_id`)
    ON DELETE CASCADE,
  CONSTRAINT `fk_student_subject_enrollments_subject`
    FOREIGN KEY (`subject_id`) REFERENCES `subjects` (`subject_id`)
    ON DELETE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
