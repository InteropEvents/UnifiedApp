import * as React from 'react';
import { Divider } from '@fluentui/react-components';
import { usePageHeaderStyles } from '../styles/Styles';

export interface IPageHeaderProps {
  title: string;
  description: any;
}

export const PageHeader: React.FunctionComponent<IPageHeaderProps> = props => {
  const styles = usePageHeaderStyles();
  return (
    <div>
      <h1>{props.title}</h1>
      <div>{props.description}</div>
      <Divider className={styles.divider} />
    </div>
  );
};
