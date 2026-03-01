package com.scriptsentries.service;

import com.scriptsentries.dto.CollabDto;
import com.scriptsentries.model.Comment;
import com.scriptsentries.model.Notification;
import com.scriptsentries.model.RiskFlag;
import com.scriptsentries.model.User;
import com.scriptsentries.repository.CommentRepository;
import com.scriptsentries.repository.NotificationRepository;
import com.scriptsentries.repository.RiskFlagRepository;
import com.scriptsentries.repository.UserRepository;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

@Service
@RequiredArgsConstructor
@Slf4j
public class CommentService {

    private final CommentRepository commentRepository;
    private final NotificationRepository notificationRepository;
    private final UserRepository userRepository;
    private final RiskFlagRepository riskFlagRepository;

    // Matches @username patterns in comment text
    private static final Pattern MENTION_PATTERN = Pattern.compile("@(\\w+)");

    @Transactional
    public CollabDto.CommentResponse addComment(CollabDto.CommentRequest request) {
        User author = userRepository.findById(request.getAuthorId())
                .orElseThrow(() -> new RuntimeException("Author not found: " + request.getAuthorId()));

        RiskFlag riskFlag = riskFlagRepository.findById(request.getRiskFlagId())
                .orElseThrow(() -> new RuntimeException("Risk flag not found: " + request.getRiskFlagId()));

        Comment comment = Comment.builder()
                .text(request.getText())
                .author(author)
                .riskFlag(riskFlag)
                .build();

        comment = commentRepository.save(comment);
        log.info("Comment saved by @{} on risk #{}", author.getUsername(), riskFlag.getId());

        // Detect @mentions and create notifications
        processMentions(request.getText(), author, riskFlag);

        return CollabDto.CommentResponse.from(comment);
    }

    public List<CollabDto.CommentResponse> getCommentsForRisk(Long riskFlagId) {
        RiskFlag riskFlag = riskFlagRepository.findById(riskFlagId)
                .orElseThrow(() -> new RuntimeException("Risk flag not found: " + riskFlagId));

        return commentRepository.findByRiskFlagOrderByCreatedAtAsc(riskFlag)
                .stream()
                .map(CollabDto.CommentResponse::from)
                .toList();
    }

    public List<CollabDto.NotificationResponse> getNotifications(Long userId) {
        User user = userRepository.findById(userId)
                .orElseThrow(() -> new RuntimeException("User not found: " + userId));

        return notificationRepository.findByRecipientOrderByCreatedAtDesc(user)
                .stream()
                .map(CollabDto.NotificationResponse::from)
                .toList();
    }

    public long getUnreadCount(Long userId) {
        User user = userRepository.findById(userId)
                .orElseThrow(() -> new RuntimeException("User not found: " + userId));
        return notificationRepository.countByRecipientAndIsReadFalse(user);
    }

    @Transactional
    public void markAllRead(Long userId) {
        User user = userRepository.findById(userId)
                .orElseThrow(() -> new RuntimeException("User not found: " + userId));

        List<Notification> unread = notificationRepository
                .findByRecipientAndIsReadFalseOrderByCreatedAtDesc(user);

        unread.forEach(n -> n.setRead(true));
        notificationRepository.saveAll(unread);
    }

    /**
     * Scans comment text for @username patterns.
     * For each valid mention found, creates a Notification for that user.
     */
    private void processMentions(String text, User author, RiskFlag riskFlag) {
        Matcher matcher = MENTION_PATTERN.matcher(text);
        List<String> notified = new ArrayList<>();

        while (matcher.find()) {
            String mentionedUsername = matcher.group(1);

            // Don't notify yourself
            if (mentionedUsername.equalsIgnoreCase(author.getUsername())) continue;

            // Already notified this user in this comment
            if (notified.contains(mentionedUsername.toLowerCase())) continue;

            userRepository.findByUsername(mentionedUsername).ifPresent(mentionedUser -> {
                String message = String.format(
                        "@%s mentioned you in a comment on risk #%d (%s): \"%s\"",
                        author.getUsername(),
                        riskFlag.getId(),
                        riskFlag.getEntityName(),
                        truncate(text, 80)
                );

                Notification notification = Notification.builder()
                        .recipient(mentionedUser)
                        .message(message)
                        .riskFlag(riskFlag)
                        .isRead(false)
                        .build();

                notificationRepository.save(notification);
                notified.add(mentionedUsername.toLowerCase());
                log.info("Notification created for @{} mentioned by @{}",
                        mentionedUsername, author.getUsername());
            });
        }
    }

    private String truncate(String text, int max) {
        return text.length() > max ? text.substring(0, max) + "..." : text;
    }
}
