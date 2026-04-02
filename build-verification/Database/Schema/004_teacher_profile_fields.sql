USE `school_management_system`;

SET @schema_name = DATABASE();

SET @middle_name_exists = (
    SELECT COUNT(*)
    FROM `information_schema`.`COLUMNS`
    WHERE `TABLE_SCHEMA` = @schema_name
      AND `TABLE_NAME` = 'teachers'
      AND `COLUMN_NAME` = 'middle_name'
);

SET @middle_name_sql = IF(
    @middle_name_exists = 0,
    'ALTER TABLE `teachers` ADD COLUMN `middle_name` VARCHAR(100) NULL AFTER `first_name`;',
    'SELECT 1;'
);
PREPARE middle_name_stmt FROM @middle_name_sql;
EXECUTE middle_name_stmt;
DEALLOCATE PREPARE middle_name_stmt;

SET @department_label_exists = (
    SELECT COUNT(*)
    FROM `information_schema`.`COLUMNS`
    WHERE `TABLE_SCHEMA` = @schema_name
      AND `TABLE_NAME` = 'teachers'
      AND `COLUMN_NAME` = 'department_label'
);

SET @department_label_sql = IF(
    @department_label_exists = 0,
    'ALTER TABLE `teachers` ADD COLUMN `department_label` VARCHAR(120) NULL AFTER `department_id`;',
    'SELECT 1;'
);
PREPARE department_label_stmt FROM @department_label_sql;
EXECUTE department_label_stmt;
DEALLOCATE PREPARE department_label_stmt;

SET @advisory_section_exists = (
    SELECT COUNT(*)
    FROM `information_schema`.`COLUMNS`
    WHERE `TABLE_SCHEMA` = @schema_name
      AND `TABLE_NAME` = 'teachers'
      AND `COLUMN_NAME` = 'advisory_section'
);

SET @advisory_section_sql = IF(
    @advisory_section_exists = 0,
    'ALTER TABLE `teachers` ADD COLUMN `advisory_section` VARCHAR(60) NULL AFTER `position_title`;',
    'SELECT 1;'
);
PREPARE advisory_section_stmt FROM @advisory_section_sql;
EXECUTE advisory_section_stmt;
DEALLOCATE PREPARE advisory_section_stmt;

SET @photo_path_exists = (
    SELECT COUNT(*)
    FROM `information_schema`.`COLUMNS`
    WHERE `TABLE_SCHEMA` = @schema_name
      AND `TABLE_NAME` = 'teachers'
      AND `COLUMN_NAME` = 'photo_path'
);

SET @photo_path_sql = IF(
    @photo_path_exists = 0,
    'ALTER TABLE `teachers` ADD COLUMN `photo_path` VARCHAR(500) NULL AFTER `advisory_section`;',
    'SELECT 1;'
);
PREPARE photo_path_stmt FROM @photo_path_sql;
EXECUTE photo_path_stmt;
DEALLOCATE PREPARE photo_path_stmt;
