USE `school_management_system`;

SET @schema_name = DATABASE();

SET @course_label_exists = (
    SELECT COUNT(*)
    FROM `information_schema`.`COLUMNS`
    WHERE `TABLE_SCHEMA` = @schema_name
      AND `TABLE_NAME` = 'subjects'
      AND `COLUMN_NAME` = 'course_label'
);

SET @course_label_sql = IF(
    @course_label_exists = 0,
    'ALTER TABLE `subjects` ADD COLUMN `course_label` VARCHAR(120) NULL AFTER `course_id`;',
    'SELECT 1;'
);
PREPARE course_label_stmt FROM @course_label_sql;
EXECUTE course_label_stmt;
DEALLOCATE PREPARE course_label_stmt;

SET @year_level_exists = (
    SELECT COUNT(*)
    FROM `information_schema`.`COLUMNS`
    WHERE `TABLE_SCHEMA` = @schema_name
      AND `TABLE_NAME` = 'subjects'
      AND `COLUMN_NAME` = 'year_level'
);

SET @year_level_sql = IF(
    @year_level_exists = 0,
    'ALTER TABLE `subjects` ADD COLUMN `year_level` VARCHAR(30) NULL AFTER `course_label`;',
    'SELECT 1;'
);
PREPARE year_level_stmt FROM @year_level_sql;
EXECUTE year_level_stmt;
DEALLOCATE PREPARE year_level_stmt;
