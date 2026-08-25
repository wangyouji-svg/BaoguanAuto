CREATE TABLE IF NOT EXISTS request_cache (
    token VARCHAR(64) PRIMARY KEY,
    trace_id VARCHAR(191) NOT NULL,
    payload LONGTEXT NOT NULL,
    created_at BIGINT NOT NULL,
    expires_at BIGINT NOT NULL,
    INDEX idx_request_cache_expires_at (expires_at),
    INDEX idx_request_cache_trace_id (trace_id)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

-- statement --
CREATE TABLE IF NOT EXISTS generated_files (
    token VARCHAR(64) PRIMARY KEY,
    trace_id VARCHAR(191) NOT NULL,
    filename VARCHAR(512) NOT NULL,
    created_at BIGINT NOT NULL,
    last_access_at BIGINT NOT NULL,
    INDEX idx_generated_files_trace_id (trace_id),
    INDEX idx_generated_files_last_access (last_access_at)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

-- statement --
CREATE TABLE IF NOT EXISTS application_logs (
    id BIGINT PRIMARY KEY AUTO_INCREMENT,
    occurred_at DATETIME(3) NOT NULL,
    level VARCHAR(16) NOT NULL,
    event_type VARCHAR(32) NOT NULL DEFAULT '',
    trace_id VARCHAR(191) NOT NULL DEFAULT '',
    method VARCHAR(16) NOT NULL DEFAULT '',
    path VARCHAR(512) NOT NULL DEFAULT '',
    status INT NULL,
    took_ms BIGINT NULL,
    contract_no VARCHAR(255) NOT NULL DEFAULT '',
    token VARCHAR(64) NOT NULL DEFAULT '',
    filename VARCHAR(512) NOT NULL DEFAULT '',
    row_count INT NULL,
    message LONGTEXT NOT NULL,
    source VARCHAR(64) NOT NULL,
    source_fingerprint CHAR(64) NOT NULL,
    created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
    UNIQUE KEY uq_application_logs_fingerprint (source_fingerprint),
    INDEX idx_application_logs_occurred (occurred_at),
    INDEX idx_application_logs_trace (trace_id),
    INDEX idx_application_logs_event_time (event_type, occurred_at),
    INDEX idx_application_logs_contract (contract_no)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

-- statement --
CREATE TABLE IF NOT EXISTS data_migration_runs (
    run_id CHAR(64) PRIMARY KEY,
    source_database VARCHAR(512) NOT NULL,
    status VARCHAR(32) NOT NULL,
    report_json LONGTEXT NOT NULL,
    started_at DATETIME(3) NOT NULL,
    completed_at DATETIME(3) NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

-- statement --
CREATE OR REPLACE VIEW generation_activity_view AS
SELECT occurred_at, trace_id, event_type, contract_no, token, filename,
       row_count, status, took_ms, message
FROM application_logs
WHERE event_type IN ('CACHE', 'BIZ', 'ERR', 'UNHANDLED');
