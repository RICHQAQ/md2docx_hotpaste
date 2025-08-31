"""Insert target selection logic."""

from ...core.types import InsertTarget
from ...infra.process import detect_active_target


class TargetSelector:
    """插入目标选择器"""
    
    def resolve_target(self, configured_target: InsertTarget) -> str:
        """
        解析实际的插入目标
        
        Args:
            configured_target: 配置的插入目标
            
        Returns:
            实际的插入目标: "word", "wps", "none"
        """
        if configured_target == "auto":
            # 自动检测当前活跃的应用
            active = detect_active_target()
            return active if active else "word"  # 默认回退到 Word
        
        elif configured_target in ("word", "wps", "none"):
            return configured_target
        
        else:
            # 未知配置，回退到默认值
            return "word"
