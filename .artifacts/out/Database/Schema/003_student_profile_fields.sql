USE `school_management_system`;

SET @schema_name = DATABASE();

SET @middle_name_exists = (
    SELECT COUNT(*)
    FROM `information_schema`.`COLUMNS`
    WHERE `TABLE_SCHEMA` = @schema_name
      AND `TABLE_NAME` = 'students'
      AND `COLUMN_NAME` = 'middle_name'
);

SET @middle_name_sql = IF(
    @middle_name_exists = 0,
    'ALTER TABLE `students` ADD COLUMN `middle_name` VARCHAR(100) NULL AFTER `first_name`;',
    'SELECT 1;'
);
PREPARE middle_name_stmt FROM @middle_name_sql;
EXECUTE middle_name_stmt;
DEALLOCATE PREPARE middle_name_stmt;

SET @photo_path_exists = (
    SELECT COUNT(*)
    FROM `information_schema`.`COLUMNS`
    WHERE `TABLE_SCHEMA` = @schema_name
      AND `TABLE_NAME` = 'students'
      AND `COLUMN_NAME` = 'photo_path'
);

SET @photo_path_sql = IF(
    @photo_path_exists = 0,
    'ALTER TABLE `students` ADD COLUMN `photo_path` VARCHAR(500) NULL AFTER `section_name`;',
    'SELECT 1;'
);
PREPARE photo_path_stmt FROM @photo_path_sql;
EXECUTE photo_path_stmt;
DEALLOCATE PREPARE photo_path_stmt;
