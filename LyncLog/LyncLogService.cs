#region License, Terms and Author(s)
//
// Lynclog, raw logging for Lync and Skype for business conversations
// Copyright (c) 2016 Philippe Raemy. All rights reserved.
//
//  Author(s):
//
//      Philippe Raemy
//
// Licensed under the Apache License, Version 2.0 (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy of the License at
//
//    http://www.apache.org/licenses/LICENSE-2.0
//
// Unless required by applicable law or agreed to in writing, software
// distributed under the License is distributed on an "AS IS" BASIS,
// WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
// See the License for the specific language governing permissions and
// limitations under the License.
//
#endregion

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using Microsoft.Lync.Model;
using Microsoft.Lync.Model.Conversation;
using MoreLinq;

namespace LyncLog
{
    internal class LyncLogServiceState {
        public LyncClient Client { get; set; }
    }

	/// <summary>
	/// https://www.stevejgordon.co.uk/using-generic-host-in-dotnet-core-console-based-microservices
	/// </summary>
	public class LyncLogService : IHostedService, IDisposable {
		private Timer _timer;
		static readonly List<ConversationContainer> ActiveConversations =
			new List<ConversationContainer>();
        static string GetRunningMethodName => new StackTrace().GetFrames()?[1].GetMethod().Name;

		public ILogger<LyncLogService> Logger { get; }
		public AppConfig AppConfig { get; }

		public LyncLogService(ILogger<LyncLogService> logger,
			IOptions<AppConfig> appConfig) {
			Logger = logger;
			AppConfig = appConfig.Value;
		}

		public Task StartAsync(CancellationToken cancellationToken) {
			Logger.LogInformation("LyncLog starting...");
			InitializeTrace();

            var lynclogServiceState = new LyncLogServiceState();
			_timer = new Timer(DoWork, lynclogServiceState, TimeSpan.Zero,
				TimeSpan.FromMilliseconds(1000));

			return Task.CompletedTask;
		}

		private void DoWork(object state) {
			// Logger.LogInformation($"Background work with text: {_appConfig.Value.TextToPrint}");
			// LyncClient client = null;
            var serviceState = (LyncLogServiceState)state;
			if (serviceState.Client == null) {
				try {
					var msgsent = false;
					while (serviceState.Client == null) {
						try {
							serviceState.Client = LyncClient.GetClient();
						} catch (ClientNotFoundException) {
							if (!msgsent) {
								Trace.TraceWarning("The Lync client is not running. Waiting... ");
								Logger.LogWarning("The Lync client is not running. Waiting... ");
								msgsent = true;
							}
							// checking that the state is active
							if (3 != ((dynamic)serviceState.Client?.InnerObject)?.State) {
								serviceState.Client = null;
							}
							Thread.Sleep(1000);
						}
					}

					Trace.TraceWarning("The Lync client is running.");
					Logger.LogWarning("The Lync client is running.");

					var handler = typeof(ConversationManager)
						.GetField("ConversationAdded", BindingFlags.NonPublic | BindingFlags.Instance)?.GetValue(serviceState.Client.ConversationManager) as Delegate;
					// check if there are already event handlers. We can assume that all events are subscribed at once, or not at all.
					if (handler?.GetInvocationList().Length > 0) {
						Console.WriteLine("InitializeTracking: ConversationManager events already subscribed. Skipping!");
					} else {
						serviceState.Client.ConversationManager.ConversationAdded -= Catcher<ConversationManagerEventArgs>(ConversationManager_ConversationAdded);
						serviceState.Client.ConversationManager.ConversationRemoved -= Catcher<ConversationManagerEventArgs>(ConversationManager_ConversationRemoved);
						serviceState.Client.ConversationManager.ConversationAdded += Catcher<ConversationManagerEventArgs>(ConversationManager_ConversationAdded);
						serviceState.Client.ConversationManager.ConversationRemoved += Catcher<ConversationManagerEventArgs>(ConversationManager_ConversationRemoved);
						Console.WriteLine("InitializeTracking: ConversationManager events subscribed.");
					}
					serviceState.Client.ConversationManager.Conversations.ForEach(InitializeTracking);

					ActiveConversations.ForEach(kvp => kvp.DumpConversation());
				} catch (Exception e) {
					Trace.TraceError(e.ToString());
				}
			}
			try {
				if (((dynamic)serviceState.Client?.InnerObject)?.State != 3) {
					Console.WriteLine("The Lync client appears to have changed status.");
					serviceState.Client = null;
				}
			} catch {
				serviceState.Client = null;
			}
			if (serviceState.Client == null) {
				Console.WriteLine("The Lync client has been lost.");
			}
		}

		public Task StopAsync(CancellationToken cancellationToken) {
			Logger.LogInformation("Stopping LyncLog.");

			_timer?.Change(Timeout.Infinite, 0);

			return Task.CompletedTask;
		}

		public void Dispose() {
			_timer?.Dispose();
		}

		private void InitializeTrace() {
            if (!Directory.Exists(AppConfig.LogFolder))
                Directory.CreateDirectory(AppConfig.LogFolder);
			var logFolder = new DirectoryInfo(AppConfig.LogFolder);
			// if (!logFolder.Exists) logFolder.Create();

			var consoleTraceListener = new ConsoleTraceListener {
				Filter = new EventTypeFilter(SourceLevels.All)
			};
			var textWriterTraceListener =
				new TextWriterTraceListener(
					new StreamWriter(Path.Combine(logFolder.FullName, $"LyncLog-{DateTime.Now:yyyyMMdd-HHmmss}.log"))) {
					TraceOutputOptions = TraceOptions.DateTime,
					Filter = new EventTypeFilter(SourceLevels.All)
				};
			Trace.Listeners.AddRange(new[] { consoleTraceListener, (TraceListener)textWriterTraceListener });
			Trace.AutoFlush = true;
			Trace.TraceWarning("Lynclog started");
		}

		private void InitializeTracking(Conversation conversation) {
            if (!Directory.Exists(AppConfig.ConversationLog))
                Directory.CreateDirectory(AppConfig.ConversationLog);
			var logFolder = new DirectoryInfo(AppConfig.ConversationLog);

			// if (conversation.State != ConversationState.Active) return;
			StoreConversation(conversation);
			var handler = typeof(Conversation).GetField("ContextDataReceived", BindingFlags.NonPublic | BindingFlags.Instance)?.GetValue(conversation) as Delegate;
			// check if there are already event handlers. We can assume that all events are subscribed at once, or not at all.
			if (handler?.GetInvocationList().Length > 0) {
				Console.WriteLine("InitializeTracking: Conversation events already subscribed. Skipping!");
				return;
			}

			conversation.ContextDataReceived += Catcher<ContextEventArgs>(Conversation_ContextDataReceived);
			conversation.ContextDataSent += Catcher<ContextEventArgs>(Conversation_ContextDataSent);
			conversation.InitialContextReceived += Catcher<InitialContextEventArgs>(Conversation_InitialContextReceived);
			conversation.InitialContextSent += Catcher<InitialContextEventArgs>(Conversation_InitialContextSent);
			conversation.ActionAvailabilityChanged += Catcher<ConversationActionAvailabilityEventArgs>(Conversation_ActionAvailabilityChanged);
			conversation.PropertyChanged += Catcher<ConversationPropertyChangedEventArgs>(Conversation_PropertyChanged);
			Console.WriteLine("InitializeTracking: Conversation events subscribed.");
		}

		private EventHandler<T> Catcher<T>(Action<object, T> action) =>
			(o, evt) => {
				try {
					action(o, evt);
				} catch (Exception ex) {
					Trace.TraceError(ex.ToString());
				}
			};

		private void Conversation_PropertyChanged(object sender, ConversationPropertyChangedEventArgs e) {
			Trace.WriteLine($"{GetRunningMethodName}:{e.Property.ToString("G")}={e.Value}");

			if ((long)e.Property == 1342701598) // IM last received timestamp
			{
				var lastTimeStamp = new DateTime(((DateTime)e.Value).Ticks, DateTimeKind.Utc).ToLocalTime();
				FindContainerOf((Conversation)sender)?.DumpConversation(lastTimeStamp);
			}
		}

		private ConversationContainer FindContainerOf(Conversation conversation) {
			if (conversation == null) return null;
			var id = conversation.Properties[ConversationProperty.Id]?.ToString();
			var r2 = conversation.Properties[ConversationProperty.Reserved2]?.ToString();

			return ActiveConversations
				.FirstOrDefault(cc => cc.Conversation.Properties[ConversationProperty.Id]?.ToString() == id
									  || cc.Conversation.Properties[ConversationProperty.Reserved2]?.ToString() == r2
				);
		}

		private void Conversation_InitialContextSent(object sender, InitialContextEventArgs e) {
			Trace.WriteLine($"{GetRunningMethodName}:{e.ApplicationData}");
		}

		private void Conversation_ActionAvailabilityChanged(object sender, ConversationActionAvailabilityEventArgs e) {
			Trace.WriteLine($"{GetRunningMethodName}:{e.Action}");
		}

		private void Conversation_InitialContextReceived(object sender, InitialContextEventArgs e) {
			Trace.WriteLine($"{GetRunningMethodName}:{e.ApplicationData}");
		}

		private void Conversation_ContextDataSent(object sender, ContextEventArgs e) {
			Trace.WriteLine($"{GetRunningMethodName}:{e.ContextData}");
		}

		private void Conversation_ContextDataReceived(object sender, ContextEventArgs e) {
			Trace.WriteLine($"Conversation_ContextDataReceived:{e.ContextData}");
		}

		private void ConversationManager_ConversationAdded(object sender, ConversationManagerEventArgs e) {
			InitializeTracking(e.Conversation);

			if (e.Conversation.Modalities[ModalityTypes.AudioVideo].State != ModalityState.Disconnected) {
				StoreConversation(e.Conversation);
			} else {
				e.Conversation.Modalities[ModalityTypes.AudioVideo].ModalityStateChanged += Program_ModalityStateChanged;
			}
		}

		private void Program_ModalityStateChanged(object sender, ModalityStateChangedEventArgs e) {
			//in this case, any state change will be from Disconnected and will therefore indicate some A/V activity
			var modality = sender as Microsoft.Lync.Model.Conversation.AudioVideo.AVModality;

			if (modality == null || FindContainerOf(modality.Conversation) != null) return;
			StoreConversation(modality.Conversation);
			modality.ModalityStateChanged -= Program_ModalityStateChanged;
		}

		private ConversationContainer StoreConversation(Conversation conversation) {
			// if(conversation.State!=ConversationState.Active)return null;
			var container = new ConversationContainer(Logger, AppConfig) {
				Conversation = conversation,
				ConversationCreated = DateTime.Now
			};
			ActiveConversations.Add(container);
			return container;
		}

		private void ConversationManager_ConversationRemoved(object sender, ConversationManagerEventArgs e) {
			// DumpConversation(e.Conversation);
			var container = FindContainerOf(e.Conversation);
			if (container != null) {
				ActiveConversations.Remove(container);
			}
		}
	}
}
