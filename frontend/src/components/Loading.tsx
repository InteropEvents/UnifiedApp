import * as React from 'react';
import { useLoadingStyles } from '../styles/Styles';

export interface ILoadingProps extends MgtTemplateProps {
  message?: string;
}
export interface MgtTemplateProps<T = any> {
  template?: string;
  dataContext?: T;
}

export const Loading: React.FunctionComponent<ILoadingProps> = (props: ILoadingProps) => {
  const styles = useLoadingStyles();
  return (
    <div className={styles.root}>
      <div className={styles.message}>
        <span>{props.message || 'Loading...'}</span>
      </div>
    </div>
  );
};
