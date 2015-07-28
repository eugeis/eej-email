package ee.email.core.elastic;

import java.util.Date;
import java.util.List;
import java.util.concurrent.atomic.AtomicLong;

import ee.email.core.ParsedCallback;
import ee.email.model.Email;

public class IndexerOfParsedEmails implements ParsedCallback<Email> {

	private final Indexer<Email> indexer;

	protected Date maxEmailDate;

	protected AtomicLong maxId;

	public IndexerOfParsedEmails(Indexer<Email> indexer) {

		super();
		this.indexer = indexer;
		this.maxId = new AtomicLong(1);
	}

	public void setMaxId(long maxId) {
		this.maxId.set(maxId + 1);
	}

	@Override
	public void parsed(String parentReference, Email entity) {
		fillGenericIdIfNotSet(entity);
		this.indexer.index(parentReference, entity);
		checkMaxEmailDate(entity);
	}

	protected void checkMaxEmailDate(Email entity) {

		Date emailDate = entity.getDate();
		if (this.maxEmailDate == null
				|| (emailDate != null && emailDate.after(this.maxEmailDate))) {
			this.maxEmailDate = emailDate;
		}
	}

	@Override
	public void parsed(String parentReference, List<Email> entities) {
		for (Email entity : entities) {
			fillGenericIdIfNotSet(entity);
		}
		this.indexer.index(parentReference, entities);
		for (Email email : entities) {
			checkMaxEmailDate(email);
		}
	}

	private void fillGenericIdIfNotSet(Email entity) {
		if (entity.getId() == null || entity.getId().isEmpty()) {
			entity.setId(Long.toString(maxId.getAndIncrement()));
		}
	}

	public Date getMaxEmailDate() {

		return this.maxEmailDate;
	}

	public long getMaxId() {
		return maxId.get() - 1;
	}
}
