USE `school_management_system`;

SET @schema_name = DATABASE();

SET @department_label_exists = (
    SELECT COUNT(*)
    FROM `information_schema`.`COLUMNS`
    WHERE `TABLE_SCHEMA` = @schema_name
      AND `TABLE_NAME` = 'courses'
      AND `COLUMN_NAME` = 'department_label'
);

SET @department_label_sql = IF(
    @department_label_exists = 0,
    'ALTER TABLE `courses` ADD COLUMN `department_label` VARCHAR(120) NULL AFTER `department_id`;',
    'SELECT 1;'
);
PREPARE department_label_stmt FROM @department_label_sql;
EXECUTE department_label_stmt;
DEALLOCATE PREPARE department_label_stmt;

SET @units_exists = (
    SELECT COUNT(*)
    FROM `information_schema`.`COLUMNS`
    WHERE `TABLE_SCHEMA` = @schema_name
      AND `TABLE_NAME` = 'courses'
      AND `COLUMN_NAME` = 'units'
);

SET @units_sql = IF(
    @units_exists = 0,
    'ALTER TABLE `courses` ADD COLUMN `units` VARCHAR(30) NULL AFTER `department_label`;',
    'SELECT 1;'
);
PREPARE units_stmt FROM @units_sql;
EXECUTE units_stmt;
DEALLOCATE PREPARE units_stmt;
