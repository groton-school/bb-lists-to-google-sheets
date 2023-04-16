export default (...tokens: string[]) =>
  `org.groton.BbListsToGoogleSheets.${tokens.join('.')}`;
