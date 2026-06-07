-- 1. Khởi tạo Cơ sở dữ liệu
CREATE DATABASE IF NOT EXISTS `ai_tapchi` CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci;
USE `ai_tapchi`;

-- 2. Cấu hình ràng buộc hệ thống
SET FOREIGN_KEY_CHECKS = 0;

-- 3. Bảng Người dùng (Lưu trữ thông tin cá nhân, vai trò, trạng thái khóa)
CREATE TABLE IF NOT EXISTS `user` (
  `id` INT(11) NOT NULL AUTO_INCREMENT,
  `email` VARCHAR(150) NOT NULL,
  `name` VARCHAR(150) DEFAULT NULL,
  `password` VARCHAR(255) DEFAULT NULL COMMENT 'Mật khẩu đã mã hóa',
  `google_id` VARCHAR(200) DEFAULT NULL,
  `auth_type` VARCHAR(20) DEFAULT 'local' COMMENT 'local hoặc google',
  `role` VARCHAR(20) DEFAULT 'user' COMMENT 'user hoặc admin',
  `is_active` TINYINT(1) DEFAULT 1 COMMENT '1: Hoạt động, 0: Bị khóa',
  `created_at` DATETIME DEFAULT CURRENT_TIMESTAMP,
  `last_login` DATETIME DEFAULT NULL,
  `picture` VARCHAR(255) DEFAULT NULL COMMENT 'URL ảnh đại diện',
  PRIMARY KEY (`id`),
  UNIQUE KEY `idx_email` (`email`),
  INDEX `idx_role` (`role`),
  INDEX `idx_auth` (`auth_type`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- 4. Bảng Thanh toán (Lưu lịch sử nạp tiền và trạng thái duyệt)
CREATE TABLE IF NOT EXISTS `payment` (
  `id` INT(11) NOT NULL AUTO_INCREMENT,
  `user_id` INT(11) NOT NULL,
  `payment_code` VARCHAR(32) NOT NULL COMMENT 'Mã chuyển khoản duy nhất',
  `amount` INT(11) NOT NULL,
  `currency` VARCHAR(10) DEFAULT 'VND',
  `status` VARCHAR(20) DEFAULT 'pending' COMMENT 'pending hoặc paid',
  `transfer_content` VARCHAR(120) NOT NULL,
  `sepay_transaction_id` VARCHAR(80) DEFAULT NULL,
  `created_at` DATETIME DEFAULT CURRENT_TIMESTAMP,
  `paid_at` DATETIME DEFAULT NULL,
  PRIMARY KEY (`id`),
  UNIQUE KEY `idx_payment_code` (`payment_code`),
  INDEX `idx_status` (`status`),
  CONSTRAINT `fk_payment_user` FOREIGN KEY (`user_id`) REFERENCES `user` (`id`) ON DELETE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- 5. Bảng Lịch sử Tạp chí (Lưu thông tin các file PDF đã tạo)
CREATE TABLE IF NOT EXISTS `magazine_history` (
  `id` INT(11) NOT NULL AUTO_INCREMENT,
  `user_id` INT(11) NOT NULL,
  `title` VARCHAR(255) DEFAULT NULL,
  `template` VARCHAR(50) DEFAULT NULL,
  `pdf_filename` VARCHAR(255) DEFAULT NULL,
  `created_at` DATETIME DEFAULT CURRENT_TIMESTAMP,
  PRIMARY KEY (`id`),
  INDEX `idx_user_mag` (`user_id`),
  CONSTRAINT `fk_magazine_user` FOREIGN KEY (`user_id`) REFERENCES `user` (`id`) ON DELETE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- 6. Mở lại ràng buộc khóa ngoại
SET FOREIGN_KEY_CHECKS = 1;

-- 7. Khởi tạo tài khoản Admin (Admin123)
-- Lưu ý: Mật khẩu đã được hash đúng chuẩn Werkzeug/Flask
INSERT INTO `user` (`email`, `name`, `password`, `role`, `auth_type`, `is_active`) 
VALUES (
  'admin@gmail.com', 
  'Quản trị viên', 
  'scrypt:32768:8:1$hqEailOVpjgGGbxN$2dd821fd209406bd321a4f70e14cde28181e764dd5035b81995a842a83014ce831cd7fb1da2ec7d144f368a5dec6b9853a63a8103ad2a2ebee91878cfaff2b69', 
  'admin', 
  'local', 
  1
) ON DUPLICATE KEY UPDATE `role`='admin';
