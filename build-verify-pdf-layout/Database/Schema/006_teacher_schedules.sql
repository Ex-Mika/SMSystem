USE `school_management_system`;

CREATE TABLE IF NOT EXISTS `teacher_schedules` (
  `schedule_id` INT NOT NULL AUTO_INCREMENT,
  `teacher_id` INT NOT NULL,
  `day_name` VARCHAR(20) NOT NULL,
  `session_label` VARCHAR(30) NOT NULL,
  `section_name` VARCHAR(60) NULL,
  `subject_code` VARCHAR(50) NULL,
  `subject_name` VARCHAR(150) NULL,
  `room_label` VARCHAR(60) NULL,
  `created_at` DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
  `updated_at` DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  PRIMARY KEY (`schedule_id`),
  UNIQUE KEY `uq_teacher_schedules_teacher_day_session` (`teacher_id`, `day_name`, `session_label`),
  KEY `ix_teacher_schedules_day_session` (`day_name`, `session_label`),
  CONSTRAINT `fk_teacher_schedules_teacher`
    FOREIGN KEY (`teacher_id`) REFERENCES `teachers` (`teacher_id`)
    ON DELETE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
