<Subscribe xmlns="http://opcfoundation.org/webservices/XMLDA/1.0/"
   ReturnValuesOnReply="true"
   SubscriptionPingRate="1000">
 <Options   ReturnErrorText="true"
            ReturnDiagnosticInfo="true"
            ReturnItemTime="true"
            ReturnItemName="true"
            ReturnItemPath="true"
            ClientRequestHandle="XYZ"
            LocaleID="" />
  <ItemList Deadband="0"
            RequestedSamplingRate="1000"
            EnableBuffering="false" >
    <Items  ItemName="Random.Int4" ClientItemHandle="10958524" />
  </ItemList>
</Subscribe>