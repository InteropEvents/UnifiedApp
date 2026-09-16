// src/MGTInternal/components/viewTypeConverter.ts
// 独立于 mgt 的 viewTypeConverter 工具函数

export type PersonViewType = 'image' | 'oneline' | 'twolines' | 'threelines' | 'fourlines';

/**
 * 将字符串转换为 Person 组件支持的 view 类型
 * @param view 字符串类型的 view
 * @returns 'image' | 'oneline' | 'twolines' | 'threelines' | 'fourlines'
 */
export function viewTypeConverter(view: string): PersonViewType {
  if (
    view === 'image' ||
    view === 'oneline' ||
    view === 'twolines' ||
    view === 'threelines' ||
    view === 'fourlines'
  ) {
    return view;
  }
  return 'oneline'; // 默认值
}
