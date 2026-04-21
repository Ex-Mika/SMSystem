USE `school_management_system`;

START TRANSACTION;

SET @seed_password_hash = '100000:AQIDBAUGBwgJCgsMDQ4PEA==:7gQDaNbD2TJ9Tv/U3z+oOOw+byXCRpvOoV5EbjMrc1w=';

INSERT INTO `departments` (
  `department_code`,
  `department_name`
)
VALUES (
  'CCS',
  'College of Computing Studies'
)
ON DUPLICATE KEY UPDATE
  `department_name` = VALUES(`department_name`);

INSERT INTO `courses` (
  `course_code`,
  `course_name`,
  `department_id`
)
SELECT
  'BSCS',
  'Bachelor of Science in Computer Science',
  `department_id`
FROM `departments`
WHERE `department_code` = 'CCS'
ON DUPLICATE KEY UPDATE
  `course_name` = VALUES(`course_name`),
  `department_id` = VALUES(`department_id`);

INSERT INTO `users` (
  `role_key`,
  `username`,
  `email`,
  `password_hash`,
  `is_active`
)
VALUES (
  'admin',
  'System Administrator',
  'admin@prmsu.edu.ph',
  @seed_password_hash,
  1
)
ON DUPLICATE KEY UPDATE
  `role_key` = VALUES(`role_key`),
  `username` = VALUES(`username`),
  `password_hash` = VALUES(`password_hash`),
  `is_active` = VALUES(`is_active`);

INSERT INTO `administrators` (
  `user_id`,
  `admin_code`,
  `first_name`,
  `middle_name`,
  `last_name`,
  `role_title`,
  `photo_path`
)
SELECT
  `user_id`,
  'ADM-0001',
  'System',
  NULL,
  'Administrator',
  'System Administrator',
  NULL
FROM `users`
WHERE `email` = 'admin@prmsu.edu.ph'
ON DUPLICATE KEY UPDATE
  `admin_code` = VALUES(`admin_code`),
  `first_name` = VALUES(`first_name`),
  `middle_name` = VALUES(`middle_name`),
  `last_name` = VALUES(`last_name`),
  `role_title` = VALUES(`role_title`),
  `photo_path` = VALUES(`photo_path`);

INSERT INTO `users` (
  `role_key`,
  `username`,
  `email`,
  `password_hash`,
  `is_active`
)
VALUES (
  'teacher',
  'Maria Santos',
  'maria.santos@prmsu.edu.ph',
  @seed_password_hash,
  1
)
ON DUPLICATE KEY UPDATE
  `role_key` = VALUES(`role_key`),
  `username` = VALUES(`username`),
  `password_hash` = VALUES(`password_hash`),
  `is_active` = VALUES(`is_active`);

INSERT INTO `teachers` (
  `user_id`,
  `employee_number`,
  `first_name`,
  `last_name`,
  `department_id`,
  `position_title`
)
SELECT
  u.`user_id`,
  'T-2026-0001',
  'Maria',
  'Santos',
  d.`department_id`,
  'Instructor'
FROM `users` u
CROSS JOIN `departments` d
WHERE u.`email` = 'maria.santos@prmsu.edu.ph'
  AND d.`department_code` = 'CCS'
ON DUPLICATE KEY UPDATE
  `employee_number` = VALUES(`employee_number`),
  `first_name` = VALUES(`first_name`),
  `last_name` = VALUES(`last_name`),
  `department_id` = VALUES(`department_id`),
  `position_title` = VALUES(`position_title`);

INSERT INTO `users` (
  `role_key`,
  `username`,
  `email`,
  `password_hash`,
  `is_active`
)
VALUES (
  'student',
  'Juan Dela Cruz',
  'juan.delacruz@student.prmsu.edu.ph',
  @seed_password_hash,
  1
)
ON DUPLICATE KEY UPDATE
  `role_key` = VALUES(`role_key`),
  `username` = VALUES(`username`),
  `password_hash` = VALUES(`password_hash`),
  `is_active` = VALUES(`is_active`);

INSERT INTO `students` (
  `user_id`,
  `student_number`,
  `first_name`,
  `last_name`,
  `course_id`,
  `year_level`,
  `section_name`
)
SELECT
  u.`user_id`,
  '2026-0001',
  'Juan',
  'Dela Cruz',
  c.`course_id`,
  1,
  'A'
FROM `users` u
CROSS JOIN `courses` c
WHERE u.`email` = 'juan.delacruz@student.prmsu.edu.ph'
  AND c.`course_code` = 'BSCS'
ON DUPLICATE KEY UPDATE
  `student_number` = VALUES(`student_number`),
  `first_name` = VALUES(`first_name`),
  `last_name` = VALUES(`last_name`),
  `course_id` = VALUES(`course_id`),
  `year_level` = VALUES(`year_level`),
  `section_name` = VALUES(`section_name`);

COMMIT;
