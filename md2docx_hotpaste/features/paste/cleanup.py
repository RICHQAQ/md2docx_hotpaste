"""File cleanup strategies."""

import os
import time

from ...core.constants import CLEANUP_DELAY
from ...infra.logging import log


class FileCleanupManager:
    """文件清理管理器"""
    
    def cleanup_if_needed(
        self,
        file_path: str,
        keep_file: bool,
        insert_success: bool,
        target: str
    ) -> None:
        """
        根据配置和插入结果决定是否清理文件
        
        Args:
            file_path: 文件路径
            keep_file: 是否保留文件
            insert_success: 插入是否成功
            target: 插入目标
        """
        if keep_file:
            # 配置为保留文件，不清理
            return
        
        try:
            # 只有在成功插入或目标为 none 时才清理
            if insert_success or target == "none":
                # 稍等一下再清理，确保文件不被占用
                time.sleep(CLEANUP_DELAY)
                
                if os.path.exists(file_path):
                    os.remove(file_path)
                    log(f"Cleaned up temporary file: {file_path}")
                    
        except Exception as e:
            log(f"Failed to cleanup file {file_path}: {e}")
            # 清理失败不影响主流程
