import * as React from 'react';
import styles from './RegistrationForm.module.scss';
import type { IRegistrationFormProps } from './IRegistrationFormProps';
import { escape } from '@microsoft/sp-lodash-subset';
import BasicForm from './ReactUnderstanding/BasicForm';
import UseStateHooks from './ReactUnderstanding/UserStateHooks';
import UseEffectHooks from './ReactUnderstanding/UseEffect';

/* export default class RegistrationForm extends React.Component<IRegistrationFormProps> {
  public render(): React.ReactElement<IRegistrationFormProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <section className={`${styles.registrationForm} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.welcome}>
          <img alt="" src={isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png')} className={styles.welcomeImage} />
          <h2>Well done, {escape(userDisplayName)}!</h2>
          <div>{environmentMessage}</div>
          <div>Web part property value: <strong>{escape(description)}</strong></div>
        </div>
        <div>
          <h3>Welcome to SharePoint Framework!</h3>
          <p>
            The SharePoint Framework (SPFx) is a extensibility model for Microsoft Viva, Microsoft Teams and SharePoint. It&#39;s the easiest way to extend Microsoft 365 with automatic Single Sign On, automatic hosting and industry standard tooling.
          </p>
          <h4>Learn more about SPFx development:</h4>
          <ul className={styles.links}>
            <li><a href="https://aka.ms/spfx" target="_blank" rel="noreferrer">SharePoint Framework Overview</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-graph" target="_blank" rel="noreferrer">Use Microsoft Graph in your solution</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-teams" target="_blank" rel="noreferrer">Build for Microsoft Teams using SharePoint Framework</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-viva" target="_blank" rel="noreferrer">Build for Microsoft Viva Connections using SharePoint Framework</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-store" target="_blank" rel="noreferrer">Publish SharePoint Framework applications to the marketplace</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-api" target="_blank" rel="noreferrer">SharePoint Framework API reference</a></li>
            <li><a href="https://aka.ms/m365pnp" target="_blank" rel="noreferrer">Microsoft 365 Developer Community</a></li>
          </ul>
        </div>
      </section>
    );
  }
} */
const RegistrationForm:React.FC<IRegistrationFormProps> = (props) =>
{
//const cannot be reassigned
  const num1=30;
  console.log(num1);
 //let and var are the same , but usually use let not var
  let num2=90;
  console.log(num2);

  num2=100;
  console.log(num2);

  let name="Yosra"
  console.log(name);

  let permission=true;
  console.log(typeof(permission));

let fruits=["Apples","oranges","mango"];
console.log(typeof(fruits));
console.log(fruits);
//to print each item of the array for each for big data set
fruits.forEach((item)=>{
  console.log(item);
})


//for loop
console.log("for loop for fruits");
for(let i=0;i<fruits.length;i++)
{
  console.log(i + " : Fruits is :" + fruits[i] );
//console.log();
 //console.log(fruits[i]);
}

 //while for large dataset, but for each is most popular
let j=0;
while(j<fruits.length)
{
  console.log(fruits[j]);
  j++;
}

let d=0;
do{
   console.log(fruits[d]);
   d++;
}
while(d<fruits.length)

  return(
    <> 
    <p className={styles.h4}>I am learning SPFx</p>
    {/* Child component */}
    <BasicForm/>

  {/* //map 2 */}
  {fruits.map((item)=>{
  return <p key={item}>{item}</p>
  })}

  {/*Hooks calling*/}
  <UseStateHooks/>
  <UseEffectHooks/>




    </>
  )
}
export default RegistrationForm;