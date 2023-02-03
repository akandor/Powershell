#################################################
# MS Teams Module INI Creator for Audiocodes SBC
# Version 1.0.3 (Build: 1.0.3.2023.01.04)
# Created by: Armin Toepper
# Date: 2022/12/30
#
# Version history at the end of file.
#################################################

# Upload directly via REST Api or create a INI file only (coming soon)

[bool] $createFileOnly = $true
# Only needed if $createFileOnly is $false
$sbcIPAddress = "[IP-Address]"
$sbcUsername = "[Username]"
$sbcPassword = "[Password]"

#################################################
#                                               #
#           Input customer data here            #
#                                               #
#################################################

$indexNo = 10
$sbcName = "SampleName"

$sigInterfaceName = "untrusted"
$mediaInterfaceName = "untrusted"

$sigTCPPort = 0
$sigUDPPort = 0
$sigTLSPort = 5067

$rtpStartPort = 7000
$rtpSessionLegs = 2000
$udpPortSpacing = 4 # Change only if SBC deault UDPPortSpacing setting was changed.

###################################################################################################################################################################################################
#
# Should be manually configured! Not fully implemented yet
#
$teamsInterface = @{IPAddress='192.168.1.100'; PrefixLength='24'; Gateway='192.168.2.1'; PrimaryDNSServerIPAddress='8.8.8.8'; SecondaryDNSServerIPAddress='8.8.4.4'; UnderlyingDevice='DeviceName'}
#[bool] $seperateMediaInterface = $true
$mediaInterface = @{IPAddress='192.168.1.101'; PrefixLength='24'; Gateway='192.168.2.1'; PrimaryDNSServerIPAddress='8.8.8.8'; SecondaryDNSServerIPAddress='8.8.4.4'; UnderlyingDevice='DeviceName'}
#
###################################################################################################################################################################################################

$publicFQDN = "sbc1.example.com"
$publicIP = "79.79.79.79"
$publicIPRTP = "78.78.78.78"

$primaryNTP = "192.168.2.1"
$secondaryNTP = "192.168.2.1"
$ntpOffset = 3600
$ntpInterfaceName = "O+M+C"

#################################################
#                                               #
#               Firewall Settings               #
#                                               #
#################################################

[bool] $setupFirewall = $false
[bool] $insertBlockRule = $false
$publicDNSServer = @("8.8.8.8","8.8.4.4")
$teamsSigEntries = @{SourceIP='52.112.0.0'; SubnetPrefix='14'; Protocol='TCP'; Description='Teams Sig 1'},
                        @{SourceIP='52.120.0.0'; SubnetPrefix='14'; Protocol='TCP'; Description='Teams Sig 2'}
$teamsMediaEntries = @{SourceIP='52.112.0.0'; SubnetPrefix='14'; Protocol='UDP'; Description='Teams Media 1'},
                        @{SourceIP='52.120.0.0'; SubnetPrefix='14'; Protocol='UDP'; Description='Teams Media 2'},
                        @{SourceIP='13.107.64.0'; SubnetPrefix='18'; Protocol='UDP'; Description='Teams Media 3'}
[bool] $useOtherSources = $false # Only if needed if other communication over the teams interface 
$otherAllowedSources = @{SourceIP='20.0.0.0'; SubnetPrefix='32'; Protocol='UDP'; Description='Example Name 1'},
                        @{SourceIP='21.0.0.0'; SubnetPrefix='32'; Protocol='UDP'; Description='Example Name 2'}

#################################################
#                                               #
#              DO NOT EDIT HERE  !!!            #
#                                               #
#################################################

$proxyIndexNo1 = $indexNo+1
$proxyIndexNo2 = $indexNo+2
$proxyIndexNo3 = $indexNo+3

$messageManSetID1 = $indexNo+1
$messageManSetID2 = $indexNo+1

$messageManipulationIndexNo1 = $indexNo+1
$messageManipulationIndexNo2 = $indexNo+2
$messageManipulationIndexNo3 = $indexNo+3
$messageManipulationIndexNo4 = $indexNo+4
$messageManipulationIndexNo5 = $indexNo+5
$messageManipulationIndexNo6 = $indexNo+6
$messageManipulationIndexNo7 = $indexNo+7

$natIndexNo1 = $indexNo+1
$natIndexNo2 = $indexNo+2
$natIndexNo3 = $indexNo+3

$rtpEndPort = $rtpStartPort + ($rtpSessionLegs * 4) - 1

$LF = "`r`n"

#################################################
#                                               #
#                 SBC INI FILE                  #
#              DO NOT EDIT HERE  !!!            #
#                                               #
#################################################

$iniHeader = @"
;****************************
;** $sbcName Ini File
;** MS Teams Module
;****************************

;Created: $((Get-Date).ToString())

"@

$iniBasics = @"
[SYSTEM Params]

NTPServerUTCOffset = $ntpOffset
NTPServerIP = '$primaryNTP'
NTPSecondaryServerIP = '$secondaryNTP'

[Voice Engine Params]

ENABLEMEDIASECURITY = 1

[ NTPInterface ]

FORMAT Index = InterfaceName;
NTPInterface 0 = "$ntpInterfaceName";

"@

$sigIFIPAddress = $teamsInterface.IPAddress
$sigIFPrefixLength = $teamsInterface.PrefixLength
$sigIFGateway = $teamsInterface.Gateway
$sigIFPrimaryDNS = $teamsInterface.PrimaryDNSServerIPAddress
$sigIFSecondaryDNS = $teamsInterface.SecondaryDNSServerIPAddress
$sigIFUnderlayingDevice = $teamsInterface.UnderlyingDevice

$interfaceTable = @"
[ InterfaceTable ]

FORMAT Index = ApplicationTypes, InterfaceMode, IPAddress, PrefixLength, Gateway, InterfaceName, PrimaryDNSServerIPAddress, SecondaryDNSServerIPAddress, UnderlyingDevice, OverwriteDynamicDNSServers;
InterfaceTable $indexNo = 4, 10, $sigIFIPAddress, $sigIFPrefixLength, $sigIFGateway, "$sigInterfaceName", $sigIFPrimaryDNS, $sigIFSecondaryDNS, "$sigIFUnderlayingDevice", 0;

"@

if($seperateMediaInterface) {

    $mediaIFIPAddress = $mediaInterface.IPAddress
    $mediaIFPrefixLength = $mediaInterface.PrefixLength
    $mediaIFGateway = $mediaInterface.Gateway
    $mediaIFPrimaryDNS = $mediaInterface.PrimaryDNSServerIPAddress
    $mediaIFSecondaryDNS = $mediaInterface.SecondaryDNSServerIPAddress
    $mediaIFUnderlayingDevice = $mediaInterface.UnderlyingDevice

    $interfaceIndexNo = $indexNo + 1

    $entry = "InterfaceTable $interfaceIndexNo = 1, 10, $sigIFIPAddress, $sigIFPrefixLength, $sigIFGateway, `"$mediaInterfaceName`", $sigIFPrimaryDNS, $sigIFSecondaryDNS, `"$sigIFUnderlayingDevice`", 0;"
    $interfaceTable = $interfaceTable + $entry + $LF
}

$interfaceTable = $interfaceTable + $LF + "[ \InterfaceTable ]"

$tlsContext = @"

[ TLSContexts ]

FORMAT Index = Name, TLSVersion, DTLSVersion, ServerCipherString, ClientCipherString, ServerCipherTLS13String, ClientCipherTLS13String, KeyExchangeGroups, RequireStrictCert, TlsRenegotiation, MiddleboxCompatMode, OcspEnable, OcspInterface, OcspServerPrimary, OcspServerSecondary, OcspServerPort, OcspDefaultResponse, UseDefaultCABundle, DHKeySize;
TLSContexts $indexNo = "Teams", 12, 0, "DEFAULT", "DEFAULT", "TLS_AES_256_GCM_SHA384:TLS_CHACHA20_POLY1305_SHA256:TLS_AES_128_GCM_SHA256", "TLS_AES_256_GCM_SHA384:TLS_CHACHA20_POLY1305_SHA256:TLS_AES_128_GCM_SHA256", "X25519:P-256:P-384:X448", 0, 1, 0, 0, "$sigInterfaceName", "0.0.0.0", "0.0.0.0", 2560, 0, 0, 2048;

[ \TLSContexts ]

"@

$ipProfile = @"
[ IpProfile ]

FORMAT Index = ProfileName, IpPreference, CodersGroupName, IsFaxUsed, JitterBufMinDelay, JitterBufOptFactor, IPDiffServ, SigIPDiffServ, RTPRedundancyDepth, CNGmode, VxxTransportType, NSEMode, IsDTMFUsed, PlayRBTone2IP, EnableEarlyMedia, ProgressIndicator2IP, EnableEchoCanceller, CopyDest2RedirectNumber, MediaSecurityBehaviour, CallLimit, DisconnectOnBrokenConnection, FirstTxDtmfOption, SecondTxDtmfOption, RxDTMFOption, EnableHold, InputGain, VoiceVolume, AddIEInSetup, SBCExtensionCodersGroupName, MediaIPVersionPreference, TranscodingMode, SBCAllowedMediaTypes, SBCAllowedAudioCodersGroupName, SBCAllowedVideoCodersGroupName, SBCAllowedCodersMode, SBCMediaSecurityBehaviour, SBCCryptoGroupName, SBCRFC2833Behavior, SBCAlternativeDTMFMethod, SBCSendMultipleDTMFMethods, SBCReceiveMultipleDTMFMethods, SBCAssertIdentity, AMDSensitivityParameterSuit, AMDSensitivityLevel, AMDMaxGreetingTime, AMDMaxPostSilenceGreetingTime, SBCDiversionMode, SBCHistoryInfoMode, EnableQSIGTunneling, SBCFaxCodersGroupName, SBCFaxBehavior, SBCFaxOfferMode, SBCFaxAnswerMode, SbcPrackMode, SBCSessionExpiresMode, SBCRemoteUpdateSupport, SBCRemoteReinviteSupport, SBCRemoteDelayedOfferSupport, SBCRemoteReferBehavior, SBCRemote3xxBehavior, SBCRemoteMultiple18xSupport, SBCRemoteEarlyMediaResponseType, SBCRemoteEarlyMediaSupport, EnableSymmetricMKI, MKISize, SBCEnforceMKISize, SBCRemoteEarlyMediaRTP, SBCRemoteSupportsRFC3960, SBCRemoteCanPlayRingback, EnableEarly183, EarlyAnswerTimeout, SBC2833DTMFPayloadType, SBCUserRegistrationTime, ResetSRTPStateUponRekey, AmdMode, SBCReliableHeldToneSource, GenerateSRTPKeys, SBCPlayHeldTone, SBCRemoteHoldFormat, SBCRemoteReplacesBehavior, SBCSDPPtimeAnswer, SBCPreferredPTime, SBCUseSilenceSupp, SBCRTPRedundancyBehavior, SBCPlayRBTToTransferee, SBCRTCPMode, SBCJitterCompensation, SBCRemoteRenegotiateOnFaxDetection, JitterBufMaxDelay, SBCUserBehindUdpNATRegistrationTime, SBCUserBehindTcpNATRegistrationTime, SBCSDPHandleRTCPAttribute, SBCRemoveCryptoLifetimeInSDP, SBCIceMode, SBCRTCPMux, SBCMediaSecurityMethod, SBCHandleXDetect, SBCRTCPFeedback, SBCRemoteRepresentationMode, SBCKeepVIAHeaders, SBCKeepRoutingHeaders, SBCKeepUserAgentHeader, SBCRemoteMultipleEarlyDialogs, SBCRemoteMultipleAnswersMode, SBCDirectMediaTag, SBCAdaptRFC2833BWToVoiceCoderBW, CreatedByRoutingServer, UsedByRoutingServer, SBCFaxReroutingMode, SBCMaxCallDuration, SBCGenerateRTP, SBCISUPBodyHandling, SBCISUPVariant, SBCVoiceQualityEnhancement, SBCMaxOpusBW, SBCEnhancedPlc, LocalRingbackTone, LocalHeldTone, SBCGenerateNoOp, SBCRemoveUnKnownCrypto, SBCMultipleCoders, DataDiffServ, SBCMSRPReinviteUpdateSupport, SBCMSRPOfferSetupRole, SBCMSRPEmpMsg, SBCRenumberMID, SBCAllowOnlyNegotiatedPT, RTCPEncryption, SBCRemoveCSRC;
IpProfile $indexNo = "IPP_Teams", 1, "AudioCodersGroups_10", 0, 10, 10, 46, 24, 0, 0, 2, 0, 0, 0, 0, -1, 1, 0, 0, -1, 1, 4, -1, 1, 1, 0, 0, "", "", 0, 0, "", "", "", 0, 1, "", 1, 0, 0, 0, 0, 0, 8, 300, 400, 0, 0, 0, "", 0, 0, 1, 3, 0, 0, 1, 0, 3, 2, 1, 0, 1, 0, 0, 0, 1, 0, 1, 0, 0, 0, 0, 0, 0, 1, 0, 0, 3, 1, 0, 0, 0, 0, 0, 1, 0, 0, 300, -1, -1, 0, 0, 0, 0, 0, 0, 0, -1, -1, -1, -1, -1, 0, "", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, -1, -1, 0, 0, 0, 0, 1, 2, 0, 0, 0, 2, 0;

[ \IpProfile ]

"@

$mediaRealm = @"
[ CpMediaRealm ]

FORMAT Index = MediaRealmName, IPv4IF, IPv6IF, RemoteIPv4IF, RemoteIPv6IF, PortRangeStart, MediaSessionLeg, PortRangeEnd, TCPPortRangeStart, TCPPortRangeEnd, IsDefault, QoeProfile, BWProfile, TopologyLocation, UsedByRoutingServer;
CpMediaRealm $indexNo = "MR_Teams", "$mediaInterfaceName", "", "", "", $rtpStartPort, $rtpSessionLegs, $rtpEndPort, 0, 0, 0, "", "", 1, 0;

[ \CpMediaRealm ]

"@

$dialPlan = @"
[ DialPlan ]

FORMAT Index = Name, PrefixCaseSensitivity;
DialPlan $indexNo = "DP_Teams", 0;

[ \DialPlan ]

"@

$sipInterface = @"
[ SIPInterface ]

FORMAT Index = InterfaceName, NetworkInterface, SCTPSecondaryNetworkInterface, ApplicationType, UDPPort, TCPPort, TLSPort, SCTPPort, AdditionalUDPPorts, AdditionalUDPPortsMode, SRDName, MessagePolicyName, TLSContext, TLSMutualAuthentication, TCPKeepaliveEnable, ClassificationFailureResponseType, PreClassificationManSet, EncapsulatingProtocol, MediaRealm, SBCDirectMedia, BlockUnRegUsers, MaxNumOfRegUsers, EnableUnAuthenticatedRegistrations, UsedByRoutingServer, TopologyLocation, PreParsingManSetName, AdmissionProfile, CallSetupRulesSetId;
SIPInterface $indexNo = "SIF_Teams", "$sigInterfaceName", "", 2, $sigTCPPort, $sigUDPPort, $sigTLSPort, 0, "", 0, "DefaultSRD", "", "Teams", 1, 1, 0, 3, 0, "MR_Teams", 0, -1, -1, -1, 0, 1, "", "", -1;

[ \SIPInterface ]

"@

$proxySet = @"
[ ProxySet ]

FORMAT Index = ProxyName, EnableProxyKeepAlive, ProxyKeepAliveTime, ProxyLoadBalancingMethod, IsProxyHotSwap, SRDName, ClassificationInput, TLSContextName, ProxyRedundancyMode, DNSResolveMethod, KeepAliveFailureResp, GWIPv4SIPInterfaceName, SBCIPv4SIPInterfaceName, GWIPv6SIPInterfaceName, SBCIPv6SIPInterfaceName, MinActiveServersLB, SuccessDetectionRetries, SuccessDetectionInterval, FailureDetectionRetransmissions, AcceptDHCPProxyList;
ProxySet $indexNo = "PS_Teams", 1, 60, 2, 1, "DefaultSRD", 0, "Teams", -1, 1, "", "", "SIF_Teams", "", "", 1, 1, 10, -1, 0;

[ \ProxySet ]

"@

$ipGroup = @"
[ IPGroup ]

FORMAT Index = Type, Name, ProxySetName, VoiceAIConnector, SIPGroupName, ContactUser, SipReRoutingMode, AlwaysUseRouteTable, SRDName, MediaRealm, InternalMediaRealm, ClassifyByProxySet, ProfileName, MaxNumOfRegUsers, InboundManSet, OutboundManSet, RegistrationMode, AuthenticationMode, MethodList, SBCServerAuthType, OAuthHTTPService, EnableSBCClientForking, SourceUriInput, DestUriInput, ContactName, UsernameAsClient, PasswordAsClient, UsernameAsServer, PasswordAsServer, UUIFormat, QOEProfile, BWProfile, AlwaysUseSourceAddr, MsgManUserDef1, MsgManUserDef2, SIPConnect, SBCPSAPMode, DTLSContext, CreatedByRoutingServer, UsedByRoutingServer, SBCOperationMode, SBCRouteUsingRequestURIPort, SBCKeepOriginalCallID, TopologyLocation, SBCDialPlanName, CallSetupRulesSetId, TeamsRegistrationMode, Tags, SBCUserStickiness, UserUDPPortAssignment, AdmissionProfile, ProxyKeepAliveUsingIPG, SBCAltRouteReasonsSetName, TeamsLocalMediaOptimization, TeamsLocalMOInitialBehavior, SIPSourceHostName, TeamsDirectRoutingMode, TeamsLocalMOSite, UserVoiceQualityReport, ValidateSourceIP, MeteringRemoteType, DedicatedConnectionMode;
IPGroup $indexNo = 0, "IPG_Teams", "PS_Teams", "", "$publicFQDN", "", -1, 0, "DefaultSRD", "MR_Teams", "", 0, "IPP_Teams", -1, $messageManSetID1, $messageManSetID2, 0, 0, "", -1, "", 0, -1, -1, "$publicFQDN", "", "", "", "", 0, "", "", 1, "", "", 0, 0, "default", 0, 0, -1, 0, 0, 1, "DP_Teams", -1, 0, "", 0, 0, "", 1, "", 0, 0, "", 0, "", 0, 0, 0, 0;

[ \IPGroup ]

"@

$proxyIp = @"
[ ProxyIp ]

FORMAT Index = ProxySetId, ProxyIpIndex, IpAddress, TransportType, Priority, Weight;
ProxyIp 0 = "$indexNo", 0, "sip.pstnhub.microsoft.com:5061", 2, 1, 1;
ProxyIp 1 = "$indexNo", 1, "sip2.pstnhub.microsoft.com:5061", 2, 2, 1;
ProxyIp 2 = "$indexNo", 2, "sip3.pstnhub.microsoft.com:5061", 2, 3, 1;

[ \ProxyIp ]

"@

$conditionTable = @"
[ ConditionTable ]

FORMAT Index = Name, Condition;
ConditionTable $indexNo = "Teams-Contact", "Header.Contact.URL.Host contains 'pstnhub.microsoft.com'";

[ \ConditionTable ]

"@

$ip2ipRouting = @"
[ IP2IPRouting ]

FORMAT Index = RouteName, RoutingPolicyName, SrcIPGroupName, SrcUsernamePrefix, SrcHost, DestUsernamePrefix, DestHost, RequestType, MessageConditionName, ReRouteIPGroupName, Trigger, CallSetupRulesSetId, DestType, DestIPGroupName, DestSIPInterfaceName, DestAddress, DestPort, DestTransportType, AltRouteOptions, GroupPolicy, CostGroup, DestTags, ModifiedDestUserName, SrcTags, IPGroupSetName, RoutingTagName, InternalAction;
IP2IPRouting 0 = "Options termination", "Default_SBCRoutingPolicy", "Any", "*", "*", "*", "*", 6, "", "Any", 0, -1, 13, "", "", "", 0, -1, 0, 0, "", "", "", "", "", "default", "Reply(Response='200')";
IP2IPRouting 1 = "REFER Re-routing", "Default_SBCRoutingPolicy", "Any", "*", "*", "*", "*", 0, "", "IPG_Teams", 2, -1, 2, "IPG_Teams", "", "", 0, -1, 0, 0, "", "", "", "", "", "default", "";

[ \IP2IPRouting ]

"@

$classification = @"
[ Classification ]

FORMAT Index = ClassificationName, MessageConditionName, SRDName, SrcSIPInterfaceName, SrcAddress, SrcPort, SrcTransportType, SrcUsernamePrefix, SrcHost, DestUsernamePrefix, DestHost, ActionType, SrcIPGroupName, DestRoutingPolicy, IpProfileName, IPGroupSelection, IpGroupTagName;
Classification $indexNo = "Teams", "Teams-Contact", "DefaultSRD", "SIF_Teams", "52.*.*.*", 0, 2, "*", "*", "*", "$publicFQDN", 1, "IPG_Teams", "", "", 0, "default";

[ \Classification ]

"@

$messageManipulations = @"
[ MessageManipulations ]

FORMAT Index = ManipulationName, ManSetID, MessageType, Condition, ActionSubject, ActionType, ActionValue, RowRole;
MessageManipulations $messageManipulationIndexNo1 = "Add Cause to History-Info", $messageManSetID1, "Invite.Request", "Header.History-Info.1 regex (<.*)(user=phone)(>)(.*)", "Header.History-Info.1", 2, "$1+$2+'?Reason=SIP%3Bcause%3D302'+$3+$4", 0;
MessageManipulations $messageManipulationIndexNo2 = "remPrivWhenNotAnon", $messageManSetID1, "Any.Request", "header.from.url !contains 'anonymous'", "header.privacy", 1, "", 0;
MessageManipulations $messageManipulationIndexNo3 = "Reject Cause", $messageManSetID1, "Any.Response", "Header.Request-Uri.MethodType=='480' OR Header.Request-Uri.MethodType=='503' OR Header.Request-Uri.MethodType=='603'", "Header.Request-Uri.MethodType", 2, "'486'", 0;
MessageManipulations $messageManipulationIndexNo4 = "Change R-URI User", $messageManSetID2, "Reinvite.Request", "", "Header.Request-URI.URL.User", 2, "Header.To.URL.User", 0;
MessageManipulations $messageManipulationIndexNo5 = "Change RecvOnly to Inactive", $messageManSetID2, "Reinvite.Request", "Param.Message.SDP.RTPMode == 'recvonly'", "Param.Message.SDP.RTPMode", 2, "'inactive'", 0;
MessageManipulations $messageManipulationIndexNo6 = "no c with zeroes to teams", $messageManSetID2, "Any", "Body.sdp regex '(.*)(c=IN IP4 0.0.0.0)(.*)'", "body.sdp", 2, "$1+'c=IN IP4 '+param.Message.SDP.OriginAddress+$3", 0;
MessageManipulations $messageManipulationIndexNo7 = "try ringing", $messageManSetID2, "Invite.Response.100", "", "Header.Request-URI.MethodType", 2, "'180'", 0;

[ \MessageManipulations ]

"@

$natTranslation = @"
[ NATTranslation ]

FORMAT Index = SrcIPInterfaceName, SourceIPAddress, RemoteInterfaceName, TargetIpMode, TargetIPAddress, SourceStartPort, SourceEndPort, TargetStartPort, TargetEndPort;
NATTranslation $natIndexNo1 = "$sigInterfaceName", "", "", 0, "$publicIP", "$sigTLSPort", "$sigTLSPort", "$sigTLSPort", "$sigTLSPort";
NATTranslation $natIndexNo2 = "$mediaInterfaceName", "", "", 0, "$publicIPRTP", "$rtpStartPort", "$rtpEndPort", "$rtpStartPort", "$rtpEndPort";

[ \NATTranslation ]

"@

$audioCoders = @"

[ AudioCoders ]

FORMAT Index = AudioCodersGroupId, AudioCodersIndex, Name, pTime, rate, PayloadType, Sce, CoderSpecific;
AudioCoders 10 = "AudioCodersGroups_10", 0, 35, 2, 19, 103, 0, "";

[ \AudioCoders ]
"@

#################################################
#                                               #
#             ALL ABOUT FIREWALLING             #
#              DO NOT EDIT HERE  !!!            #
#                                               #
#################################################

$firewall = @"
[ ACCESSLIST ]

FORMAT Index = Source_IP, Source_Port, PrefixLen, Start_Port, End_Port, Protocol, Use_Specific_Interface, Interface_ID, Packet_Size, Byte_Rate, Byte_Burst, Allow_type_enum, Description;

"@

$firewallIndexNo = $indexNo

foreach ($dnsserver in $publicDNSServer)
{
    $firewallIndexNo += 1
    $entry = "ACCESSLIST $firewallIndexNo = `"$dnsserver`", 0, 32, 53, 53, `"Any`", 1, `"$sigInterfaceName`", 0, 0, 0, 0, `"Public DNS`";"
    $firewall = $firewall + $entry + $LF
}

if($primaryNTP.Length -gt 0) {
    $firewallIndexNo += 1
    $entry = "ACCESSLIST $firewallIndexNo = `"$primaryNTP`", 0, 32, 123, 123, `"UDP`", 1, `"$ntpInterfaceName`", 0, 0, 0, 0, `"Primary NTP Server`";"
    $firewall = $firewall + $entry + $LF
}

if($secondaryNTP.Length -gt 0) {
    $firewallIndexNo += 1
    $entry = "ACCESSLIST $firewallIndexNo = `"$secondaryNTP`", 0, 32, 123, 123, `"UDP`", 1, `"$ntpInterfaceName`", 0, 0, 0, 0, `"Secondary NTP Server`";"
    $firewall = $firewall + $entry + $LF
}

foreach ($teamsSigserver in $teamsSigEntries)
{
    $firewallIndexNo += 1
    $entrySourceIP = $teamsSigserver.SourceIP
    $entrySubnetPrefix = $teamsSigserver.SubnetPrefix
    $entryProtocol = $teamsSigserver.Protocol
    $entryDescription = $teamsSigserver.Description
    $entry = "ACCESSLIST $firewallIndexNo = `"$entrySourceIP`", 0, $entrySubnetPrefix, $sigTLSPort, $sigTLSPort, `"$entryProtocol`", 1, `"$sigInterfaceName`", 0, 0, 0, 0, `"$entryDescription`";"
    $firewall = $firewall + $entry + $LF
}

foreach ($teamsMediaserver in $teamsMediaEntries)
{
    $firewallIndexNo += 1
    $entrySourceIP = $teamsMediaserver.SourceIP
    $entrySubnetPrefix = $teamsMediaserver.SubnetPrefix
    $entryProtocol = $teamsMediaserver.Protocol
    $entryDescription = $teamsMediaserver.Description
    $entry = "ACCESSLIST $firewallIndexNo = `"$entrySourceIP`", 0, $entrySubnetPrefix, $rtpStartPort, $rtpEndPort, `"$entryProtocol`", 1, `"$mediaInterfaceName`", 0, 0, 0, 0, `"$entryDescription`";"
    $firewall = $firewall + $entry + $LF
}

if($useOtherSources)
{
    foreach ($otherSources in $otherAllowedSources)
    {
        $firewallIndexNo += 1
        $entrySourceIP = $otherSources.SourceIP
        $entrySubnetPrefix = $otherSources.SubnetPrefix
        $entryProtocol = $otherSources.Protocol
        $entryDescription = $otherSources.Description
        $entry = "ACCESSLIST $firewallIndexNo = `"$entrySourceIP`", 0, $entrySubnetPrefix, 0, 65535, `"$entryProtocol`", 1, `"$sigInterfaceName`", 0, 0, 0, 0, `"$entryDescription`";"
        $firewall = $firewall + $entry + $LF
    }
}

if($insertBlockRule) {
    # Last Rule is block the rest
    $firewallIndexNo += 1
    $firewallBlockRule = "ACCESSLIST $firewallIndexNo = `"0.0.0.0`", 0, 0, 0, 65535, `"Any`", 1, `"$sigInterfaceName`", 0, 0, 0, 1, `"Block traffic`";"
    $firewall = $firewall + $firewallBlockRule + $LF
}

$firewall = $firewall + $LF + "[ \ACCESSLIST ]"

#################################################
#                                               #
#                 FILE OUTPUT                   #
#              DO NOT EDIT HERE  !!!            #
#                                               #
#################################################

$inifile = $iniHeader,
            $iniBasics,
            #$interfaceTable,
            $tlsContext,
            $ipProfile,
            $mediaRealm,
            $dialPlan,
            $sipInterface,
            $proxySet,
            $ipGroup,
            $proxyIp,
            $conditionTable,
            $ip2ipRouting,
            $classification,
            $messageManipulations,
            $natTranslation,
            $audioCoders

if($setupFirewall) {
    $inifile = $inifile,
                $firewall
}

try {
    $inifile | Out-File -FilePath "$PSScriptRoot\TeamsModule_$sbcName.ini" -encoding utf8
    Write-Host "INI File successfully created.";
}
catch {
    Write-Host "Error while writing INI File!";
}

#################################################
#                                               #
#                 FILE UPLOAD                   #
#              DO NOT EDIT HERE  !!!            #
#                                               #
#################################################

# COMING SOON!!!
if(-not ($createFileOnly)) {
$URLIncremental = "http://{0}/api/v1/files/ini/incremental" ` -f $sbcIPAddress
# REST API Authentication
$authHash = [Convert]::ToBase64String( ` [Text.Encoding]::ASCII.GetBytes( ` ("{0}:{1}" -f $sbcUsername,$sbcPassword)))

# INI File Body
$boundary = [System.Guid]::NewGuid().ToString(); 
$bodyLines = (
    "--$boundary",
("Content-Disposition: form-data; name=`"file`";" + `
" filename=`"file.txt`""),
"Content-Type: application/octet-stream$LF", $iniHeader,
"--$boundary--$LF"
) -join $LF

$uploadINI = Invoke-RestMethod -Uri $URLIncremental -Method Put ` -Headers @{Authorization=("Basic {0}" -f $authHash)} ` -ContentType "multipart/form-data; boundary=$boundary" ` -Body $bodyLines
$uploadINI | ConvertTo-Json
}


#################################################
#                                               #
#                  END OF SCRIPT                #
#                                               #
#################################################
#################################################
#################################################
#                                               #
#                 VERSION HISTORY               #
#                                               #
#################################################
#
# Build 1.0.3.2023.01.04 - Armin Toepper
# - Missing Message Manipluation 'Reject Reason'
# added
#
#------------------------------------------------
#
# Build 1.0.2.2023.01.02 - Armin Toepper
# - Fixed some bugs
# - Created INI File tested
#
#------------------------------------------------
#
# Build 1.0.1.2022.12.30 - Armin Toepper
# - NTP Server and Firewallrules added
# - Added InterfaceTable but not activated
#
#------------------------------------------------
#
# Build 1.0.2022.12.30 - Armin Toepper
# - Powershell Script created
#
#################################################