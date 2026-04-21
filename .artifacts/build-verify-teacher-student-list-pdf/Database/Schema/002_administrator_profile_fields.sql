USE `school_management_system`;

SET @schema_name = DATABASE();

SET @middle_name_exists = (
    SELECT COUNT(*)
    FROM `information_schema`.`COLUMNS`
    WHERE `TABLE_SCHEMA` = @schema_name
      AND `TABLE_NAME` = 'administrators'
      AND `COLUMN_NAME` = 'middle_name'
);

SET @middle_name_sql = IF(
    @middle_name_exists = 0,
    'ALTER TABLE `administrators` ADD COLUMN `middle_name` VARCHAR(100) NULL AFTER `first_name`;',
    'SELECT 1;'
);
PREPARE middle_name_stmt FROM @middle_name_sql;
EXECUTE middle_name_stmt;
DEALLOCATE PREPARE middle_name_stmt;

SET @role_title_exists = (
    SELECT COUNT(*)
    FROM `information_schema`.`COLUMNS`
    WHERE `TABLE_SCHEMA` = @schema_name
      AND `TABLE_NAME` = 'administrators'
      AND `COLUMN_NAME` = 'role_title'
);

SET @role_title_sql = IF(
    @role_title_exists = 0,
    'ALTER TABLE `administrators` ADD COLUMN `role_title` VARCHAR(100) NULL AFTER `last_name`;',
    'SELECT 1;'
);
PREPARE role_title_stmt FROM @role_title_sql;
EXECUTE role_title_stmt;
DEALLOCATE PREPARE role_title_stmt;

SET @photo_path_exists = (
    SELECT COUNT(*)
    FROM `information_schema`.`COLUMNS`
    WHERE `TABLE_SCHEMA` = @schema_name
      AND `TABLE_NAME` = 'administrators'
      AND `COLUMN_NAME` = 'photo_path'
);

SET @photo_path_sql = IF(
    @photo_path_exists = 0,
    'ALTER TABLE `administrators` ADD COLUMN `photo_path` VARCHAR(500) NULL AFTER `role_title`;',
    'SELECT 1;'
);
PREPARE photo_path_stmt FROM @photo_path_sql;
EXECUTE photo_path_stmt;
DEALLOCATE PREPARE photo_path_stmt;
