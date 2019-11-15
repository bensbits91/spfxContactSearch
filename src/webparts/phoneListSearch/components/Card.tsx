import * as React from 'react';
import { PersonaCoin } from 'office-ui-fabric-react/lib/Persona';
import { Link } from 'office-ui-fabric-react/lib/Link';
import styles from './PhoneListSearch.module.scss';




/* export  */interface ICardProps {
   item?: any;
   searchTerms: string;
   size?: string;
   show_department: boolean;
   show_division: boolean;
   show_organization: boolean;
}

/* export  */interface ICardState {

}


export default class Card extends React.Component<ICardProps, ICardState> {

   constructor(props) {
      super(props);
      this.state = {

      };
   }

   public render() {
      const searchTerms = this.props.searchTerms;
      let highlightHits = (str) => {
         for (let term of searchTerms) {
            const searchTermRegex = new RegExp(term.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&'), "ig");
            const searchTermHighlighted = '<span style="background-color:yellow;">$&</span>';
            str = str.replace(searchTermRegex, searchTermHighlighted);
         }
         return str;
      };

      return (
         <div
            key={this.props.item.Id}
            className={this.props.size == 'large' ? styles.contactItem : [styles.contactItem, styles.small].join(' ')}
            data-item-id={this.props.item.Id}
         >
            <div className={styles.contactItemImg}>
               <Link href={"https://delve-gcc.office.com/?p=" + this.props.item.Email + "&v=work"} target="about:blank">
                  <PersonaCoin
                     text={this.props.item.FirstName ? this.props.item.FirstName + ' ' + this.props.item.Title : this.props.item.Title}
                     coinSize={this.props.size == 'large' ? 100 : 50}
                     showInitialsUntilImageLoads={true}
                  />
               </Link>
            </div>
            <div className={styles.contactItemDetails}>
               <div className={styles.padBottom}>
                  <Link href={"https://delve-gcc.office.com/?p=" + this.props.item.Email + "&v=work"} target="about:blank">
                     <div className={[styles.contactItemFullName, styles.contactItemFieldBody].join(' ')}
                        dangerouslySetInnerHTML={{
                           __html: highlightHits(this.props.item.FirstName ? this.props.item.FirstName + ' ' + this.props.item.Title : this.props.item.Title)
                        }}
                     />
                  </Link>
                  {this.props.item.JobTitle
                     ? <div className={styles.contactItemFieldBody}
                        dangerouslySetInnerHTML={{ __html: highlightHits(this.props.item.JobTitle) }} />
                     : ''
                  }
                  {this.props.item.WorkPhone || this.props.item.CellPhone
                     ? < div className={styles.contactItemFieldBody}>
                        {this.props.item.WorkPhone
                           ? <span className={styles.contactItemFieldBody_span}>W: {this.props.item.WorkPhone}</span>
                           : ''
                        }
                        {this.props.item.CellPhone
                           ? <span className={styles.contactItemFieldBody_span}>C: {this.props.item.CellPhone}</span>
                           : ''
                        }
                     </div>
                     : ''
                  }
                  {this.props.item.Email
                     ? <div className={styles.contactItemFieldBody}>
                        <a href={'mailto:' + this.props.item.Email}>
                           {this.props.item.Email}
                        </a>
                     </div>
                     : ''
                  }
               </div>
               <div className={styles.padBottom}>
                  {this.props.item.Organization && this.props.show_organization
                     ? <div className={styles.contactItemFieldBody}
                        dangerouslySetInnerHTML={{ __html: highlightHits(this.props.item.Organization) }} />
                     : ''
                  }
                  {this.props.item.Company && this.props.show_department
                     ? <div>
                        <span className={styles.contactItemFieldLabel}>Department: </span>
                        <span className={styles.contactItemFieldBody}
                           dangerouslySetInnerHTML={{ __html: highlightHits(this.props.item.Company) }} />
                     </div>
                     : ''
                  }
                  {this.props.item.Division && this.props.show_division
                     ? <div>
                        <span className={styles.contactItemFieldLabel}>Division: </span>
                        <span className={styles.contactItemFieldBody}
                           dangerouslySetInnerHTML={{ __html: highlightHits(this.props.item.Division) }} />
                     </div>
                     : ''}
                  {this.props.item.Program
                     ? <div>
                        <span className={styles.contactItemFieldLabel}>Program: </span>
                        <span className={styles.contactItemFieldBody}
                           dangerouslySetInnerHTML={{ __html: highlightHits(this.props.item.Program) }} />
                     </div>
                     : ''}
               </div>
               {this.props.item.WorkAddress
                  ? <div className={styles.contactItemFieldBody}>{this.props.item.WorkAddress}</div>
                  : ''
               }
            </div>
         </div >
      );
   }

}

